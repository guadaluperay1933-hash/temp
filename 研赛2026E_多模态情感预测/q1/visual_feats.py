"""视觉模态：视频帧 → MediaPipe FaceLandmarker 的 52 个 ARKit 风格 blendshape → 35 维 AU 特征（OpenFace 2.0 布局）。

35 维 = 17 个 AU 强度（0~5）+ 18 个 AU 出现标志（0/1），顺序与 OpenFace 2.0 的 AUxx_r / AUxx_c 输出一致：
    强度  AU01 AU02 AU04 AU05 AU06 AU07 AU09 AU10 AU12 AU14 AU15 AU17 AU20 AU23 AU25 AU26 AU45
    出现  AU01 AU02 AU04 AU05 AU06 AU07 AU09 AU10 AU12 AU14 AU15 AU17 AU20 AU23 AU25 AU26 AU28 AU45
强度 = 5 × s_AU，s_AU ∈ [0,1] 为对应 blendshape 分数（左右成对者取平均，映射见 AU_BLENDSHAPES）；
出现 = 1[5·s_AU ≥ θ]，θ = presence_threshold（默认 1.0，即 s ≥ 0.2）。AU28（吸唇）只有出现标志。

为什么用 MediaPipe 而不是 OpenFace：OpenFace 2.0 需要自行编译 C++，Windows/无管理员环境难以复现；
MediaPipe 可 pip 安装、模型文件 3.7 MB（Apache-2.0，q1/models/face_landmarker.task），CPU 上逐帧实时。
blendshape 是按 FACS 语义命名的表情系数，与 AU 一一对应关系明确，因此可以得到与附件2 同维度同语义排列的特征。
注意：数值尺度与 OpenFace 的 AU 回归值并不相同（不同模型），论文中需说明"布局对齐、数值不可直接互换"。

未检出人脸的帧 → 整行为 0（与 common.data 的"整行为 0 = 不可用"语义一致），并统计人脸检出率。

小脸重试（crop_retry）：FaceLandmarker 的人脸检测器（BlazeFace 短距模型，输入 128×128）会把 16:9 画面按长边
缩放，远景/小脸在检测器输入里只有十几个像素而漏检。若整帧处理的检出率 < crop_retry_below，则对整段视频
再用"居中正方形裁剪"（边长 = min(H, W)）重跑一遍，取检出率更高的一遍（整段视频只用一种输入，避免逐帧混用两种
输入导致 blendshape 数值口径不一致），并记录所用的一遍。裁剪不改变时间戳，blendshape 是表情系数、与坐标无关。
"""
from __future__ import annotations

import ctypes
import os
import sys
from pathlib import Path

import numpy as np

AU_INTENSITY = ["AU01", "AU02", "AU04", "AU05", "AU06", "AU07", "AU09", "AU10", "AU12", "AU14", "AU15",
                "AU17", "AU20", "AU23", "AU25", "AU26", "AU45"]
AU_PRESENCE = ["AU01", "AU02", "AU04", "AU05", "AU06", "AU07", "AU09", "AU10", "AU12", "AU14", "AU15",
               "AU17", "AU20", "AU23", "AU25", "AU26", "AU28", "AU45"]
VISION_DIM = len(AU_INTENSITY) + len(AU_PRESENCE)
VISION_FEATURE_NAMES = [f"{a}_r" for a in AU_INTENSITY] + [f"{a}_c" for a in AU_PRESENCE]
assert VISION_DIM == 35

# AU → blendshape（多个时取平均）。AU25 另有专门公式，见 blendshapes_to_au()。
AU_BLENDSHAPES = {
    "AU01": ("browInnerUp",),                               # 内眉上扬
    "AU02": ("browOuterUpLeft", "browOuterUpRight"),        # 外眉上扬
    "AU04": ("browDownLeft", "browDownRight"),              # 皱眉（眉下压）
    "AU05": ("eyeWideLeft", "eyeWideRight"),                # 上睑提升（睁大眼）
    "AU06": ("cheekSquintLeft", "cheekSquintRight"),        # 脸颊上提
    "AU07": ("eyeSquintLeft", "eyeSquintRight"),            # 眼睑收紧
    "AU09": ("noseSneerLeft", "noseSneerRight"),            # 皱鼻
    "AU10": ("mouthUpperUpLeft", "mouthUpperUpRight"),      # 上唇提升
    "AU12": ("mouthSmileLeft", "mouthSmileRight"),          # 嘴角上扬
    "AU14": ("mouthDimpleLeft", "mouthDimpleRight"),        # 酒窝（嘴角内收）
    "AU15": ("mouthFrownLeft", "mouthFrownRight"),          # 嘴角下压
    "AU17": ("mouthShrugLower",),                           # 下巴上抬
    "AU20": ("mouthStretchLeft", "mouthStretchRight"),      # 嘴唇横向拉伸
    "AU23": ("mouthPressLeft", "mouthPressRight"),          # 嘴唇收紧/抿
    "AU26": ("jawOpen",),                                   # 下颌下降
    "AU28": ("mouthRollLower", "mouthRollUpper"),           # 吸唇
    "AU45": ("eyeBlinkLeft", "eyeBlinkRight"),              # 眨眼
}

DEFAULT_VISION_CFG = {
    "model_path": "q1/models/face_landmarker.task", "sample_fps": None, "max_side": None,
    "presence_threshold": 1.0, "au25_gain": 2.0, "min_face_detection_confidence": 0.5,
    "min_face_presence_confidence": 0.5, "min_tracking_confidence": 0.5, "gl_lib_dir": None,
    "crop_retry": True, "crop_retry_below": 0.8,
}


def blendshapes_to_au(bs: dict, presence_threshold: float = 1.0, au25_gain: float = 2.0) -> np.ndarray:
    """一帧的 blendshape 分数字典 → 35 维 AU 行。

    AU25（双唇分开）：s_25 = clip(g · max(jawOpen − mouthClose, mean(mouthLowerDownL/R)), 0, 1)，g = au25_gain。
      理由：ARKit 的 mouthClose 定义为"下颌张开时双唇仍闭合"的矫正量，所以 jawOpen − mouthClose 才是"下颌张开且唇也张开"
      的程度；不下颌也能靠下唇下拉让双唇分开，因此再与 mouthLowerDown 取大。不用 1 − mouthClose：自然闭嘴时
      mouthClose≈0，会把闭唇误判为"完全张开"。g=2：唇缝在下颌略张开时就已可见，而 AU26 直接用 jawOpen 表示下颌开度。
    """
    def s(names):
        return float(np.mean([bs.get(n, 0.0) for n in names]))

    sc = {au: s(names) for au, names in AU_BLENDSHAPES.items()}
    lips = max(bs.get("jawOpen", 0.0) - bs.get("mouthClose", 0.0),
               0.5 * (bs.get("mouthLowerDownLeft", 0.0) + bs.get("mouthLowerDownRight", 0.0)))
    sc["AU25"] = float(np.clip(au25_gain * lips, 0.0, 1.0))
    inten = np.array([5.0 * np.clip(sc[a], 0.0, 1.0) for a in AU_INTENSITY], np.float32)
    pres = np.array([1.0 if 5.0 * np.clip(sc[a], 0.0, 1.0) >= presence_threshold else 0.0 for a in AU_PRESENCE],
                    np.float32)
    return np.concatenate([inten, pres])


def preload_gl_libs(lib_dir: str | None) -> list[str]:
    """MediaPipe 的 libmediapipe.so 在 Linux 上动态链接 libEGL.so.1 / libGLESv2.so.2（即使只用 CPU 推理）。
    无图形库的服务器可 `apt-get install libegl1 libgles2`；无 root 时把这几个 .so 放进一个目录，
    在配置 vision.gl_lib_dir 指定它 —— 这里用 RTLD_GLOBAL 预加载，动态链接器会按 soname 复用已加载的库。"""
    loaded = []
    if not lib_dir or not sys.platform.startswith("linux"):
        return loaded
    d = Path(lib_dir)
    for name in ("libGLdispatch.so.0", "libEGL.so.1", "libGLESv2.so.2"):
        p = d / name
        if p.exists():
            ctypes.CDLL(str(p), mode=ctypes.RTLD_GLOBAL)
            loaded.append(str(p))
    return loaded


class FaceAUExtractor:
    """每个视频新建一个 VIDEO 模式的 FaceLandmarker（VIDEO 模式要求时间戳单调递增且会跨帧跟踪，
    不同视频之间不能共用跟踪状态）。"""

    def __init__(self, cfg: dict | None = None, root: str | Path | None = None):
        c = dict(DEFAULT_VISION_CFG)
        c.update({k: v for k, v in (cfg or {}).items() if k in DEFAULT_VISION_CFG})
        self.cfg = c
        mp_path = Path(c["model_path"])
        if not mp_path.is_absolute() and root is not None:
            mp_path = Path(root) / mp_path
        if not mp_path.exists():
            raise FileNotFoundError(f"找不到 FaceLandmarker 模型文件：{mp_path}")
        self.model_path = mp_path
        gl_dir = c["gl_lib_dir"] or os.environ.get("Q1_GL_LIB_DIR")
        self.preloaded = preload_gl_libs(gl_dir)
        try:
            import mediapipe as mp
            from mediapipe.tasks import python as mpt
            from mediapipe.tasks.python import vision as mpv
        except Exception as e:  # noqa: BLE001
            raise RuntimeError(f"无法导入 mediapipe tasks API：{e}") from e
        self._mp, self._mpt, self._mpv = mp, mpt, mpv
        self.mp_version = getattr(mp, "__version__", "?")
        # 立即建一次，尽早暴露 libEGL 缺失等环境问题
        lm = self._new_landmarker()
        lm.close()

    def _new_landmarker(self):
        mpt, mpv = self._mpt, self._mpv
        opts = mpv.FaceLandmarkerOptions(
            base_options=mpt.BaseOptions(model_asset_path=str(self.model_path)),
            running_mode=mpv.RunningMode.VIDEO, num_faces=1, output_face_blendshapes=True,
            min_face_detection_confidence=float(self.cfg["min_face_detection_confidence"]),
            min_face_presence_confidence=float(self.cfg["min_face_presence_confidence"]),
            min_tracking_confidence=float(self.cfg["min_tracking_confidence"]))
        try:
            return mpv.FaceLandmarker.create_from_options(opts)
        except OSError as e:
            raise RuntimeError(
                f"MediaPipe 加载失败（{e}）。Linux 无图形库时请安装 libegl1 libgles2，"
                "或在 configs/q1.yaml 的 vision.gl_lib_dir 指定含 libEGL.so.1/libGLESv2.so.2 的目录") from e

    def info(self) -> dict:
        return {"tool": "MediaPipe FaceLandmarker (tasks API, RunningMode.VIDEO)", "mediapipe": self.mp_version,
                "model_file": self.model_path.name, "model_bytes": self.model_path.stat().st_size,
                "num_faces": 1, "output_face_blendshapes": True, **{k: self.cfg[k] for k in (
                    "sample_fps", "max_side", "presence_threshold", "au25_gain", "min_face_detection_confidence",
                    "min_face_presence_confidence", "min_tracking_confidence", "crop_retry", "crop_retry_below")}}

    def process_video(self, path: str | Path, backend: str = "cv2") -> dict:
        """→ dict(feats (F,35) float32, times (F,), frame_idx (F,), face (F,) bool, fps, n_decoded, detect_rate, pass)。

        F = 实际处理的帧数（每帧或按 sample_fps 抽帧）；n_decoded = 最后处理帧号 + 1（逐帧处理时即实际可解码帧数）。
        pass = 'full'（整帧）或 'center_square'（小脸重试，见模块说明）。
        """
        res = self._run(path, backend, crop=None)
        res["pass"] = "full"
        res["detect_rate_full"] = res["detect_rate"]
        res["detect_rate_crop"] = None
        if self.cfg.get("crop_retry") and res["detect_rate"] < float(self.cfg.get("crop_retry_below", 0.8)):
            alt = self._run(path, backend, crop="center_square")
            res["detect_rate_crop"] = alt["detect_rate"]
            if alt["detect_rate"] > res["detect_rate"]:
                alt["pass"] = "center_square"
                alt["detect_rate_full"], alt["detect_rate_crop"] = res["detect_rate"], alt["detect_rate"]
                res = alt
        return res

    @staticmethod
    def _crop(rgb: np.ndarray, crop: str | None) -> np.ndarray:
        if crop != "center_square":
            return rgb
        h, w = rgb.shape[:2]
        s = min(h, w)
        y0, x0 = (h - s) // 2, (w - s) // 2
        return rgb[y0:y0 + s, x0:x0 + s]

    def _run(self, path: str | Path, backend: str, crop: str | None) -> dict:
        from .media import iter_frames, probe_video

        mp = self._mp
        fps = probe_video(path)["fps"] or 25.0
        lm = self._new_landmarker()
        feats, times, fidx, face = [], [], [], []
        last_ts = -1
        n_decoded = 0
        try:
            for idx, t, rgb in iter_frames(path, target_fps=self.cfg["sample_fps"], max_side=self.cfg["max_side"],
                                           backend=backend):
                n_decoded = idx + 1
                ts = int(round(t * 1000.0))
                if ts <= last_ts:  # VIDEO 模式要求毫秒时间戳严格递增
                    ts = last_ts + 1
                last_ts = ts
                img = mp.Image(image_format=mp.ImageFormat.SRGB, data=np.ascontiguousarray(self._crop(rgb, crop)))
                res = lm.detect_for_video(img, ts)
                if res.face_blendshapes:
                    bs = {c.category_name: float(c.score) for c in res.face_blendshapes[0]}
                    row = blendshapes_to_au(bs, float(self.cfg["presence_threshold"]), float(self.cfg["au25_gain"]))
                    face.append(True)
                else:
                    row = np.zeros(VISION_DIM, np.float32)
                    face.append(False)
                feats.append(row)
                times.append(t)
                fidx.append(idx)
        finally:
            lm.close()
        F = len(feats)
        face_arr = np.array(face, dtype=bool)
        return {
            "feats": np.stack(feats).astype(np.float32) if F else np.zeros((0, VISION_DIM), np.float32),
            "times": np.array(times, np.float64), "frame_idx": np.array(fidx, np.int64), "face": face_arr,
            "fps": float(fps), "n_decoded": int(n_decoded), "n_processed": F,
            "detect_rate": float(face_arr.mean()) if F else 0.0,
        }
