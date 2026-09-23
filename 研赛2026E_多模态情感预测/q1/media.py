"""音视频读取的底层工具：ffmpeg 定位、视频探测、音频解码、逐帧读取、文件校验和。

接口约定（问题3 直接 import，签名不要改）：
    probe_video(path)            -> dict(duration, fps, n_frames, width, height, has_audio, audio_sr, audio_channels, ...)
    load_audio(path, sr=16000)   -> float32 单声道波形 (n,)；文件没有音轨时返回长度 0 的数组
    iter_frames(path, step=1, target_fps=None) -> 逐个产出 (frame_idx, t_sec, rgb)
    read_frame_at(path, t_sec)   -> rgb uint8 (H, W, 3)
    file_md5(path)               -> 32 位十六进制串

为什么不用 ffprobe：imageio-ffmpeg 只带 ffmpeg 一个可执行文件。容器时长、音轨采样率/声道
从 `ffmpeg -i` 的 stderr 解析；帧率、帧数、分辨率用 OpenCV 读（两者不一致时以实际解码为准，
见 visual_feats 里的实际解码帧数）。

时间约定：视频第 n 帧的时间戳 t_n = n / fps（显示时刻，秒）；音频第 j 个采样点的时间 j / sr。
两条流都从 0 开始计时（容器 start 偏移记录在 probe 结果里，但不参与计算，MOSEI 片段均为 0）。
"""
from __future__ import annotations

import hashlib
import re
import shutil
import subprocess
from functools import lru_cache
from pathlib import Path

import numpy as np

# ----------------------------------------------------------------------------- ffmpeg

_FFMPEG_OVERRIDE: str | None = None


def set_ffmpeg(path: str | None) -> None:
    """配置文件里指定了 ffmpeg 路径时调用（优先级最高）。"""
    global _FFMPEG_OVERRIDE
    _FFMPEG_OVERRIDE = path or None
    find_ffmpeg.cache_clear()


@lru_cache(maxsize=1)
def find_ffmpeg() -> str:
    """查找顺序：set_ffmpeg 指定 → imageio-ffmpeg 自带二进制 → 系统 PATH 上的 ffmpeg。"""
    if _FFMPEG_OVERRIDE:
        return _FFMPEG_OVERRIDE
    try:
        import imageio_ffmpeg

        exe = imageio_ffmpeg.get_ffmpeg_exe()
        if exe and Path(exe).exists():
            return exe
    except Exception:  # noqa: BLE001 - 没装 imageio-ffmpeg 就退回系统 ffmpeg
        pass
    exe = shutil.which("ffmpeg")
    if exe:
        return exe
    raise FileNotFoundError("找不到 ffmpeg：请 pip install imageio-ffmpeg，或安装系统 ffmpeg 并加入 PATH")


def ffmpeg_version() -> str:
    try:
        out = subprocess.run([find_ffmpeg(), "-hide_banner", "-version"], capture_output=True, text=True, timeout=30)
        return out.stdout.splitlines()[0].strip() if out.stdout else "?"
    except Exception as e:  # noqa: BLE001
        return f"unknown ({e})"


# ----------------------------------------------------------------------------- 探测

_DUR_RE = re.compile(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)(?:,\s*start:\s*(-?\d+(?:\.\d+)?))?")
_VSTREAM_RE = re.compile(r"Stream #\d+:\d+.*?: Video: ([^,\s]+).*?(\d{2,5})x(\d{2,5})")
_FPS_RE = re.compile(r"(\d+(?:\.\d+)?)\s*fps")
_ASTREAM_RE = re.compile(r"Stream #\d+:\d+.*?: Audio: ([^,\s]+)[^\n]*?(\d+)\s*Hz,\s*([^,\n]+)")


def _channels_from_layout(layout: str) -> int | None:
    layout = layout.strip().lower()
    table = {"mono": 1, "stereo": 2, "2.1": 3, "quad": 4, "4.0": 4, "5.0": 5, "5.1": 6, "5.1(side)": 6,
             "6.1": 7, "7.1": 8, "downmix": 2}
    if layout in table:
        return table[layout]
    m = re.match(r"(\d+)\s*channels?", layout)
    return int(m.group(1)) if m else None


def _ffmpeg_info(path: str | Path) -> dict:
    """解析 `ffmpeg -i <file>` 的 stderr（没有输出文件，ffmpeg 会以非零码退出，这是正常的）。"""
    res = subprocess.run([find_ffmpeg(), "-hide_banner", "-i", str(path)], capture_output=True, text=True,
                         errors="replace", timeout=60)
    err = res.stderr
    info: dict = {"duration": None, "start": 0.0, "has_audio": False, "audio_sr": None, "audio_channels": None,
                  "audio_codec": None, "video_codec": None, "fps_ffmpeg": None, "width_ffmpeg": None,
                  "height_ffmpeg": None}
    m = _DUR_RE.search(err)
    if m:
        h, mi, s = int(m.group(1)), int(m.group(2)), float(m.group(3))
        info["duration"] = h * 3600 + mi * 60 + s
        if m.group(4) is not None:
            info["start"] = float(m.group(4))
    for line in err.splitlines():
        if ": Video:" in line and info["video_codec"] is None:
            mv = _VSTREAM_RE.search(line)
            if mv:
                info["video_codec"] = mv.group(1)
                info["width_ffmpeg"], info["height_ffmpeg"] = int(mv.group(2)), int(mv.group(3))
            mf = _FPS_RE.search(line)
            if mf:
                info["fps_ffmpeg"] = float(mf.group(1))
        elif ": Audio:" in line and not info["has_audio"]:
            info["has_audio"] = True
            ma = _ASTREAM_RE.search(line)
            if ma:
                info["audio_codec"] = ma.group(1)
                info["audio_sr"] = int(ma.group(2))
                info["audio_channels"] = _channels_from_layout(ma.group(3))
    return info


def probe_video(path: str | Path) -> dict:
    """视频基本信息。

    duration：容器时长（秒，ffmpeg 解析）；解析不到时用 n_frames / fps。
    fps / n_frames / width / height：OpenCV 读取（n_frames 来自容器元数据，可能与实际可解码帧数差 1~2 帧）。
    has_audio / audio_sr / audio_channels：原始音轨信息（重采样前）。
    """
    import cv2

    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    info = _ffmpeg_info(path)
    cap = cv2.VideoCapture(str(path))
    fps = n_frames = width = height = None
    if cap.isOpened():
        fps = float(cap.get(cv2.CAP_PROP_FPS)) or None
        n_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) or None
        width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH)) or None
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT)) or None
    cap.release()
    fps = fps or info["fps_ffmpeg"]
    width = width or info["width_ffmpeg"]
    height = height or info["height_ffmpeg"]
    duration = info["duration"]
    if duration is None and fps and n_frames:
        duration = n_frames / fps
    if n_frames is None and fps and duration:
        n_frames = int(round(duration * fps))
    return {
        "duration": float(duration) if duration is not None else None,
        "fps": float(fps) if fps else None,
        "n_frames": int(n_frames) if n_frames else None,
        "width": width, "height": height,
        "has_audio": bool(info["has_audio"]),
        "audio_sr": info["audio_sr"], "audio_channels": info["audio_channels"],
        "audio_codec": info["audio_codec"], "video_codec": info["video_codec"],
        "start": info["start"],
    }


# ----------------------------------------------------------------------------- 音频

def load_audio(path: str | Path, sr: int = 16000) -> np.ndarray:
    """用 ffmpeg 管道解码为 float32 单声道、采样率 sr 的波形（多声道取平均：ffmpeg -ac 1）。

    文件没有音轨 → 返回长度 0 的数组（调用方据此记为"语音不可用"），其它解码失败 → RuntimeError。
    """
    cmd = [find_ffmpeg(), "-hide_banner", "-nostdin", "-loglevel", "error", "-i", str(path),
           "-vn", "-ac", "1", "-ar", str(int(sr)), "-f", "f32le", "-acodec", "pcm_f32le", "pipe:1"]
    res = subprocess.run(cmd, capture_output=True, timeout=600)
    if res.returncode != 0:
        msg = res.stderr.decode("utf-8", errors="replace")
        if "does not contain any stream" in msg or "Output file #0 does not contain" in msg or \
                "matches no streams" in msg:
            return np.zeros(0, dtype=np.float32)
        raise RuntimeError(f"ffmpeg 解码音频失败：{path}\n{msg[-800:]}")
    y = np.frombuffer(res.stdout, dtype="<f4").astype(np.float32, copy=True)
    y[~np.isfinite(y)] = 0.0
    return y


# ----------------------------------------------------------------------------- 视频帧

def _select_frame(idx: int, fps: float, step: int, target_fps: float | None) -> bool:
    """按固定步长或目标帧率抽帧：目标帧率 r 下，第 idx 帧被选中 ⇔ floor(idx·r/fps) 相对上一帧增加。"""
    if target_fps and fps and target_fps < fps:
        return idx == 0 or int(np.floor(idx * target_fps / fps)) != int(np.floor((idx - 1) * target_fps / fps))
    return idx % max(1, int(step)) == 0


def _resize_max_side(rgb: np.ndarray, max_side: int | None) -> np.ndarray:
    if not max_side:
        return rgb
    import cv2

    h, w = rgb.shape[:2]
    s = max(h, w)
    if s <= max_side:
        return rgb
    k = max_side / s
    return cv2.resize(rgb, (int(round(w * k)), int(round(h * k))), interpolation=cv2.INTER_AREA)


def _iter_frames_ffmpeg(path: Path, fps: float, width: int, height: int):
    """OpenCV 打不开时的备用解码：ffmpeg 输出 rgb24 原始帧（-vsync passthrough 保证不补帧、不丢帧）。"""
    cmd = [find_ffmpeg(), "-hide_banner", "-nostdin", "-loglevel", "error", "-i", str(path), "-an",
           "-vsync", "passthrough", "-f", "rawvideo", "-pix_fmt", "rgb24", "pipe:1"]
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    nbytes = width * height * 3
    idx = 0
    try:
        while True:
            buf = proc.stdout.read(nbytes)
            if len(buf) < nbytes:
                break
            yield idx, np.frombuffer(buf, np.uint8).reshape(height, width, 3)
            idx += 1
    finally:
        proc.stdout.close()
        proc.wait()


def iter_frames(path: str | Path, step: int = 1, target_fps: float | None = None, max_side: int | None = None,
                backend: str = "cv2"):
    """逐帧解码，产出 (frame_idx, t_sec, rgb)。

    frame_idx 是原视频中的帧序号（从 0 起，抽帧时不连续），t_sec = frame_idx / fps。
    step：每 step 帧取一帧；target_fps：按目标帧率均匀抽帧（优先于 step）；max_side：长边缩放上限。
    backend='cv2' 打不开时自动退回 ffmpeg 管道解码。
    """
    import cv2

    path = Path(path)
    cap = cv2.VideoCapture(str(path)) if backend == "cv2" else None
    if cap is not None and cap.isOpened():
        fps = float(cap.get(cv2.CAP_PROP_FPS)) or None
        if not fps:
            fps = probe_video(path)["fps"] or 25.0
        idx = 0
        try:
            while True:
                ok, bgr = cap.read()
                if not ok:
                    break
                if _select_frame(idx, fps, step, target_fps):
                    rgb = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
                    yield idx, idx / fps, _resize_max_side(rgb, max_side)
                idx += 1
        finally:
            cap.release()
        return
    if cap is not None:
        cap.release()
    info = probe_video(path)
    if not (info["fps"] and info["width"] and info["height"]):
        raise RuntimeError(f"无法解码视频：{path}")
    for idx, rgb in _iter_frames_ffmpeg(path, info["fps"], info["width"], info["height"]):
        if _select_frame(idx, info["fps"], step, target_fps):
            yield idx, idx / info["fps"], _resize_max_side(rgb, max_side)


def read_frame_at(path: str | Path, t_sec: float) -> np.ndarray:
    """读取时刻 t_sec 处的帧（帧号 round(t·fps)，越界时截到最后一帧）→ RGB uint8 (H, W, 3)。

    先用 OpenCV 按帧号定位；失败时用 ffmpeg -ss 精确定位再取一帧。
    """
    import cv2

    path = Path(path)
    cap = cv2.VideoCapture(str(path))
    try:
        if cap.isOpened():
            fps = float(cap.get(cv2.CAP_PROP_FPS)) or 25.0
            n = int(cap.get(cv2.CAP_PROP_FRAME_COUNT)) or 0
            idx = int(round(max(0.0, float(t_sec)) * fps))
            if n > 0:
                idx = min(idx, n - 1)
            cap.set(cv2.CAP_PROP_POS_FRAMES, idx)
            ok, bgr = cap.read()
            if not ok and idx > 0:  # 元数据帧数偏大时，末帧可能读不到 → 往前退几帧
                for back in range(1, 6):
                    cap.set(cv2.CAP_PROP_POS_FRAMES, max(0, idx - back))
                    ok, bgr = cap.read()
                    if ok:
                        break
            if ok:
                return cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    finally:
        cap.release()
    info = probe_video(path)
    cmd = [find_ffmpeg(), "-hide_banner", "-nostdin", "-loglevel", "error", "-ss", f"{max(0.0, float(t_sec)):.3f}",
           "-i", str(path), "-frames:v", "1", "-f", "rawvideo", "-pix_fmt", "rgb24", "pipe:1"]
    res = subprocess.run(cmd, capture_output=True, timeout=60)
    w, h = info["width"], info["height"]
    if res.returncode != 0 or not w or not h or len(res.stdout) < w * h * 3:
        raise RuntimeError(f"无法读取 {path} 在 {t_sec:.3f}s 处的帧")
    return np.frombuffer(res.stdout[: w * h * 3], np.uint8).reshape(h, w, 3).copy()


# ----------------------------------------------------------------------------- 校验和

def file_md5(path: str | Path, chunk: int = 1 << 20) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()
