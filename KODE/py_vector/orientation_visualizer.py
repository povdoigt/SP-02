"""
=============================================================================
Orientation Visualizer (GPU / OpenGL)
=============================================================================
Real-time 3D orientation viewer driven by serial data. Renders a full
body-frame basis attached to a 3D object (rocket / cube / plane / your own
STL model). Uses pyqtgraph's OpenGL backend — orders of magnitude smoother
than matplotlib for live data.

Input formats (auto-cleaned: NULL bytes, +signs, whitespace all handled):
    - 3 floats  -> Euler angles in degrees  (roll, pitch, yaw)   [--mode euler]
    -            -> OR a direction vector     (x, y, z)            [--mode vector]
    - 4 floats  -> Quaternion (w, x, y, z)                        [--mode quat]
    - 9 floats  -> Basis: body X,Y,Z axes in world frame          [--mode basis]
                   order: bx_x,bx_y,bx_z, by_x,by_y,by_z, bz_x,bz_y,bz_z
                   (great for checking an on-board quaternion->vectors conversion)

Usage:
    python orientation_visualizer.py --port COM3 --object rocket --mode euler
    python orientation_visualizer.py --port COM3 --object cube   --mode quat
    python orientation_visualizer.py --port COM3 --mode basis
    python orientation_visualizer.py --port COM3 --stl myrocket.stl
    python orientation_visualizer.py --list-ports

Requirements:
    pip install pyqtgraph PyQt5 PyOpenGL numpy pyserial
=============================================================================
"""

import sys
import struct
import argparse
import threading
import re
import json

import numpy as np
import serial
from serial.tools import list_ports

import pyqtgraph as pg
import pyqtgraph.opengl as gl
from pyqtgraph.Qt import QtCore, QtWidgets


# =============================================================================
#  Serial reading (thread-safe, robust parser)
# =============================================================================

class SerialReader(threading.Thread):
    """Reads lines from the serial port and keeps the latest parsed sample."""

    def __init__(self, port, baudrate):
        super().__init__(daemon=True)
        self.port = port
        self.baudrate = baudrate
        self.running = True
        self.ser = None

        self._lock = threading.Lock()
        self._latest = None          # last parsed list[float]
        self.success_count = 0
        self.error_count = 0

    def get_latest(self):
        with self._lock:
            return None if self._latest is None else list(self._latest)

    def run(self):
        try:
            self.ser = serial.Serial(self.port, self.baudrate, timeout=1)
            print(f"Connected to {self.port} @ {self.baudrate} baud. Waiting for data...")
        except serial.SerialException as e:
            print(f"Serial error: {e}")
            print(f"Available ports: {[p.device for p in list_ports.comports()]}")
            return

        while self.running:
            try:
                raw = self.ser.readline()
                if not raw:
                    continue
                try:
                    line = raw.decode("utf-8")
                except UnicodeDecodeError:
                    self.error_count += 1
                    continue

                values = self._parse(line)
                if values is None:
                    continue

                with self._lock:
                    self._latest = values
                self.success_count += 1
                if self.success_count == 1:
                    print(f"First packet: {values}")

            except Exception as e:
                self.error_count += 1
                if self.error_count <= 3:
                    print(f"Read error: {e}")

        if self.ser and self.ser.is_open:
            self.ser.close()

    @staticmethod
    def _parse(line):
        """Return a list of floats (length 3 or 4) or None."""
        # Strip NULL bytes and non-printables (your STM32 stream had trailing 0x00)
        line = line.replace("\0", "")
        line = "".join(c for c in line if ord(c) >= 32 or c in "\t")
        line = line.strip().strip("[]{}")
        if not line:
            return None

        # JSON
        if line.startswith("{"):
            try:
                d = json.loads(line)
                if "w" in d:
                    return [float(d["w"]), float(d["x"]), float(d["y"]), float(d["z"])]
                return [float(d["x"]), float(d["y"]), float(d["z"])]
            except Exception:
                pass

        # Pull every number out (handles +signs, commas, spaces, tabs uniformly)
        nums = re.findall(r"[+-]?\d+\.?\d*(?:[eE][+-]?\d+)?", line)
        try:
            vals = [float(n) for n in nums]
        except ValueError:
            return None

        if len(vals) >= 9:
            return vals[:9]
        if len(vals) >= 4:
            return vals[:4]
        if len(vals) == 3:
            return vals[:3]
        return None

    def stop(self):
        self.running = False


# =============================================================================
#  Orientation math
# =============================================================================

def euler_to_matrix(roll_deg, pitch_deg, yaw_deg):
    """ZYX intrinsic: R = Rz(yaw) @ Ry(pitch) @ Rx(roll). Returns 4x4."""
    r, p, y = np.radians([roll_deg, pitch_deg, yaw_deg])
    cr, sr = np.cos(r), np.sin(r)
    cp, sp = np.cos(p), np.sin(p)
    cy, sy = np.cos(y), np.sin(y)

    Rx = np.array([[1, 0, 0], [0, cr, -sr], [0, sr, cr]])
    Ry = np.array([[cp, 0, sp], [0, 1, 0], [-sp, 0, cp]])
    Rz = np.array([[cy, -sy, 0], [sy, cy, 0], [0, 0, 1]])
    R = Rz @ Ry @ Rx

    M = np.eye(4)
    M[:3, :3] = R
    return M


def quat_to_matrix(w, x, y, z):
    """Quaternion (w,x,y,z) -> 4x4. Auto-normalizes."""
    n = np.sqrt(w * w + x * x + y * y + z * z)
    if n < 1e-9:
        return np.eye(4)
    w, x, y, z = w / n, x / n, y / n, z / n

    R = np.array([
        [1 - 2 * (y * y + z * z), 2 * (x * y - z * w),     2 * (x * z + y * w)],
        [2 * (x * y + z * w),     1 - 2 * (x * x + z * z), 2 * (y * z - x * w)],
        [2 * (x * z - y * w),     2 * (y * z + x * w),     1 - 2 * (x * x + y * y)],
    ])
    M = np.eye(4)
    M[:3, :3] = R
    return M


def vector_to_matrix(v):
    """Rotation that maps the object's +Z axis onto direction v. Returns 4x4."""
    v = np.asarray(v, dtype=float)
    n = np.linalg.norm(v)
    if n < 1e-9:
        return np.eye(4)
    target = v / n
    z = np.array([0.0, 0.0, 1.0])

    axis = np.cross(z, target)
    s = np.linalg.norm(axis)
    c = float(np.dot(z, target))

    M = np.eye(4)
    if s < 1e-9:  # parallel or anti-parallel
        if c < 0:
            M[:3, :3] = np.diag([1.0, -1.0, -1.0])  # 180° flip
        return M

    axis /= s
    K = np.array([[0, -axis[2], axis[1]],
                  [axis[2], 0, -axis[0]],
                  [-axis[1], axis[0], 0]])
    R = np.eye(3) + s * K + (1 - c) * (K @ K)  # Rodrigues
    M[:3, :3] = R
    return M


def basis_to_matrix(vals, transpose=False):
    """3 basis vectors (bx, by, bz) -> 4x4. NO re-orthonormalization.

    The 9 numbers are the body X, Y, Z axes expressed in the world frame,
    in order: bx_x, bx_y, bx_z, by_x, by_y, by_z, bz_x, bz_y, bz_z.
    These become the COLUMNS of R, so p_world = R @ p_body. If your firmware
    stores the axes as rows instead, pass transpose=True.
    """
    if vals is None or len(vals) < 9:
        return None
    bx = np.array(vals[0:3], dtype=float)
    by = np.array(vals[3:6], dtype=float)
    bz = np.array(vals[6:9], dtype=float)
    R = np.column_stack([bx, by, bz])
    if transpose:
        R = R.T
    M = np.eye(4)
    M[:3, :3] = R
    return M


def basis_diagnostics(vals, transpose=False):
    """Quality metrics for a basis sent by the firmware.

    Returns (norms, dots, det) where:
      norms = (|bx|, |by|, |bz|)         -> should all be ~1.0
      dots  = (bx.by, by.bz, bz.bx)      -> should all be ~0.0 (orthogonal)
      det   = det([bx by bz])            -> should be ~+1.0 (right-handed)
    """
    bx = np.array(vals[0:3], dtype=float)
    by = np.array(vals[3:6], dtype=float)
    bz = np.array(vals[6:9], dtype=float)
    if transpose:
        bx, by, bz = (np.array([vals[0], vals[3], vals[6]]),
                      np.array([vals[1], vals[4], vals[7]]),
                      np.array([vals[2], vals[5], vals[8]]))
    norms = (float(np.linalg.norm(bx)),
             float(np.linalg.norm(by)),
             float(np.linalg.norm(bz)))
    dots = (float(np.dot(bx, by)),
            float(np.dot(by, bz)),
            float(np.dot(bz, bx)))
    det = float(np.linalg.det(np.column_stack([bx, by, bz])))
    return norms, dots, det


# =============================================================================
#  Procedural meshes -> (vertices Nx3, faces Mx3, face_color RGBA)
# =============================================================================

def _cylinder(r0, r1, length, z0, seg=32):
    v, f = [], []
    for i in range(seg):
        a = 2 * np.pi * i / seg
        v.append([r0 * np.cos(a), r0 * np.sin(a), z0])
    for i in range(seg):
        a = 2 * np.pi * i / seg
        v.append([r1 * np.cos(a), r1 * np.sin(a), z0 + length])
    for i in range(seg):
        j = (i + 1) % seg
        b0, b1, t0, t1 = i, j, seg + i, seg + j
        f.append([b0, b1, t1])
        f.append([b0, t1, t0])
    return np.array(v, dtype=float), np.array(f, dtype=int)


def _disk(radius, z, seg=32, flip=False):
    v = [[0, 0, z]]
    for i in range(seg):
        a = 2 * np.pi * i / seg
        v.append([radius * np.cos(a), radius * np.sin(a), z])
    f = []
    for i in range(seg):
        j = (i + 1) % seg
        f.append([0, 1 + i, 1 + j] if not flip else [0, 1 + j, 1 + i])
    return np.array(v, dtype=float), np.array(f, dtype=int)


def _box(cx, cy, cz, sx, sy, sz):
    hx, hy, hz = sx / 2, sy / 2, sz / 2
    v = np.array([
        [cx - hx, cy - hy, cz - hz], [cx + hx, cy - hy, cz - hz],
        [cx + hx, cy + hy, cz - hz], [cx - hx, cy + hy, cz - hz],
        [cx - hx, cy - hy, cz + hz], [cx + hx, cy - hy, cz + hz],
        [cx + hx, cy + hy, cz + hz], [cx - hx, cy + hy, cz + hz],
    ], dtype=float)
    f = np.array([
        [0, 1, 2], [0, 2, 3], [4, 6, 5], [4, 7, 6],
        [0, 4, 5], [0, 5, 1], [1, 5, 6], [1, 6, 2],
        [2, 6, 7], [2, 7, 3], [3, 7, 4], [3, 4, 0],
    ], dtype=int)
    return v, f


def _combine(parts):
    """parts: list of (v, f, rgba). Returns merged verts, faces, per-face colors."""
    all_v, all_f, all_c = [], [], []
    offset = 0
    for v, f, rgba in parts:
        all_v.append(v)
        all_f.append(f + offset)
        all_c.append(np.tile(rgba, (len(f), 1)))
        offset += len(v)
    return (np.vstack(all_v),
            np.vstack(all_f),
            np.vstack(all_c).astype(float))


def mesh_rocket():
    R, body_len, nose_len = 0.35, 2.2, 1.0
    body_col = (0.85, 0.85, 0.90, 1.0)
    nose_col = (0.85, 0.20, 0.20, 1.0)
    fin_col = (0.20, 0.45, 0.85, 1.0)

    parts = []
    bv, bf = _cylinder(R, R, body_len, -body_len / 2)
    parts.append((bv, bf, body_col))
    dv, df = _disk(R, -body_len / 2, flip=True)
    parts.append((dv, df, body_col))
    nv, nf = _cylinder(R, 0.0, nose_len, body_len / 2)
    parts.append((nv, nf, nose_col))

    # 3 fins, 120° apart, at the base
    fin_h, fin_span, fin_t = 0.8, 0.55, 0.04
    base_z = -body_len / 2
    for k in range(3):
        ang = 2 * np.pi * k / 3
        fv, ff = _box(0, 0, base_z + fin_h / 2, fin_t, fin_span, fin_h)
        fv = fv + np.array([0, R + fin_span / 2 - 0.02, 0])  # push out radially
        ca, sa = np.cos(ang), np.sin(ang)
        rot = np.array([[ca, -sa, 0], [sa, ca, 0], [0, 0, 1]])
        fv = fv @ rot.T
        parts.append((fv, ff, fin_col))
    return _combine(parts)


def mesh_cube():
    v, f = _box(0, 0, 0, 1.4, 1.4, 1.4)
    cols = np.array([
        (0.85, 0.25, 0.25, 1), (0.85, 0.25, 0.25, 1),  # -Z
        (0.25, 0.85, 0.25, 1), (0.25, 0.85, 0.25, 1),  # +Z
        (0.25, 0.45, 0.85, 1), (0.25, 0.45, 0.85, 1),
        (0.90, 0.75, 0.20, 1), (0.90, 0.75, 0.20, 1),
        (0.65, 0.30, 0.80, 1), (0.65, 0.30, 0.80, 1),
        (0.25, 0.80, 0.80, 1), (0.25, 0.80, 0.80, 1),
    ], dtype=float)
    return v, f, cols


def mesh_plane():
    fus_col = (0.85, 0.85, 0.88, 1.0)
    wing_col = (0.25, 0.45, 0.85, 1.0)
    tail_col = (0.85, 0.25, 0.25, 1.0)
    parts = []
    # fuselage along +Z (nose forward), built from a cylinder + nose cone
    R, L = 0.18, 2.0
    bv, bf = _cylinder(R, R, L, -L / 2)
    parts.append((bv, bf, fus_col))
    nv, nf = _cylinder(R, 0.0, 0.6, L / 2)
    parts.append((nv, nf, fus_col))
    # main wings (span along X)
    wv, wf = _box(0, 0, -0.1, 2.6, 0.5, 0.05)
    parts.append((wv, wf, wing_col))
    # horizontal stabilizer (rear)
    hv, hf = _box(0, 0, -L / 2 + 0.15, 1.1, 0.35, 0.04)
    parts.append((hv, hf, tail_col))
    # vertical stabilizer (rear, up along Y)
    vv, vf = _box(0, 0.28, -L / 2 + 0.15, 0.05, 0.55, 0.35)
    parts.append((vv, vf, tail_col))
    return _combine(parts)


def load_stl(path):
    """Minimal binary-STL loader. Centers and scales the model to fit."""
    with open(path, "rb") as fp:
        head = fp.read(84)
        n = struct.unpack("<I", head[80:84])[0]
        rec = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("attr", "<u2")])
        data = np.frombuffer(fp.read(n * 50), dtype=rec, count=n)
    verts = data["v"].reshape(-1, 3).astype(float)
    faces = np.arange(len(verts)).reshape(-1, 3)
    # center + normalize size
    verts -= verts.mean(axis=0)
    scale = np.abs(verts).max()
    if scale > 0:
        verts *= (1.6 / scale)
    cols = np.tile((0.80, 0.80, 0.85, 1.0), (len(faces), 1)).astype(float)
    return verts, faces, cols


# =============================================================================
#  Viewer
# =============================================================================

class Viewer(QtWidgets.QMainWindow):
    def __init__(self, reader, mode, obj, stl_path=None, smooth=0.25, transpose=False):
        super().__init__()
        self.reader = reader
        self.mode = mode
        self.smooth = float(smooth)
        self.transpose = bool(transpose)
        self._M = np.eye(4)  # smoothed orientation matrix

        self.setWindowTitle("Orientation Visualizer (OpenGL)")
        self.resize(1100, 800)

        self.view = gl.GLViewWidget()
        self.view.setBackgroundColor(pg.mkColor(18, 18, 22))
        self.view.setCameraPosition(distance=6, elevation=18, azimuth=45)
        self.setCentralWidget(self.view)

        self._add_world_grid_and_axes()

        # --- object mesh ---
        if stl_path:
            v, f, c = load_stl(stl_path)
        elif obj == "cube":
            v, f, c = mesh_cube()
        elif obj == "plane":
            v, f, c = mesh_plane()
        else:
            v, f, c = mesh_rocket()

        md = gl.MeshData(vertexes=v, faces=f, faceColors=c)
        self.body = gl.GLMeshItem(meshdata=md, smooth=False,
                                  drawEdges=False, shader="shaded")
        self.view.addItem(self.body)

        # --- body-frame basis (rotates with the object) ---
        self.body_axes = self._make_body_axes(length=2.2)
        for it in self.body_axes:
            self.view.addItem(it)

        # status label overlay
        self.label = QtWidgets.QLabel(self.view)
        self.label.setStyleSheet(
            "color:#ddd; background:rgba(0,0,0,120); padding:6px;"
            "font-family:Consolas,monospace; font-size:12px;")
        self.label.move(10, 10)
        self.label.resize(400, 180)

        # ~60 FPS update
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self._update)
        self.timer.start(16)

    def _add_world_grid_and_axes(self):
        grid = gl.GLGridItem()
        grid.setSize(10, 10)
        grid.setSpacing(0.5, 0.5)
        grid.setColor((80, 80, 90, 120))
        self.view.addItem(grid)

        # faint world axes (X red, Y green, Z blue) at origin
        for vec, col in [((3, 0, 0), (1, 0.3, 0.3, 0.45)),
                         ((0, 3, 0), (0.3, 1, 0.3, 0.45)),
                         ((0, 0, 3), (0.3, 0.5, 1, 0.45))]:
            line = gl.GLLinePlotItem(
                pos=np.array([[0, 0, 0], vec], dtype=float),
                color=col, width=1.0, antialias=True)
            self.view.addItem(line)

    def _make_body_axes(self, length):
        """Three bright, thick axes representing the body frame."""
        specs = [((length, 0, 0), (1.0, 0.25, 0.25, 1.0)),   # body X
                 ((0, length, 0), (0.25, 1.0, 0.25, 1.0)),   # body Y
                 ((0, 0, length), (0.35, 0.55, 1.0, 1.0))]   # body Z
        items = []
        for vec, col in specs:
            it = gl.GLLinePlotItem(
                pos=np.array([[0, 0, 0], vec], dtype=float),
                color=col, width=4.0, antialias=True)
            it._base = np.array([[0, 0, 0], vec, [0, 0, 0]], dtype=float)[:2]
            items.append(it)
        return items

    def _orientation_matrix(self, vals):
        if vals is None:
            return None
        if self.mode == "basis":
            return basis_to_matrix(vals, self.transpose)
        if self.mode == "quat" and len(vals) >= 4:
            return quat_to_matrix(*vals[:4])
        if self.mode == "vector":
            return vector_to_matrix(vals[:3])
        # default: euler (degrees)
        return euler_to_matrix(*vals[:3])

    def _update(self):
        vals = self.reader.get_latest()
        target = self._orientation_matrix(vals)
        if target is not None:
            a = self.smooth
            if self.mode == "basis":
                # Show exactly what the firmware sends. Do NOT re-orthonormalize:
                # a skewed/scaled frame should look skewed so you can spot the bug.
                if a >= 1.0:
                    self._M = target.copy()
                else:
                    self._M[:3, :3] = (1 - a) * self._M[:3, :3] + a * target[:3, :3]
            else:
                # exponential smoothing + SVD re-orthonormalization (proper rotation)
                self._M[:3, :3] = (1 - a) * self._M[:3, :3] + a * target[:3, :3]
                U, _, Vt = np.linalg.svd(self._M[:3, :3])
                self._M[:3, :3] = U @ Vt

        tr = pg.Transform3D(*self._M.flatten().astype(float))
        self.body.setTransform(tr)

        # rotate the body axes by the same matrix
        R = self._M[:3, :3]
        for it in self.body_axes:
            base = it._base @ R.T
            it.setData(pos=base)

        self._update_label(vals)

    def _update_label(self, vals):
        if vals is None:
            txt = "Waiting for data..."
        else:
            if self.mode == "basis" and len(vals) >= 9:
                norms, dots, det = basis_diagnostics(vals, self.transpose)
                ok_n = all(abs(n - 1.0) < 0.02 for n in norms)
                ok_o = all(abs(d) < 0.02 for d in dots)
                ok_h = abs(det - 1.0) < 0.05
                txt = (
                    f"mode: basis (3 vectors){'  [rows]' if self.transpose else '  [cols]'}\n"
                    f"X=({vals[0]:+.3f},{vals[1]:+.3f},{vals[2]:+.3f})\n"
                    f"Y=({vals[3]:+.3f},{vals[4]:+.3f},{vals[5]:+.3f})\n"
                    f"Z=({vals[6]:+.3f},{vals[7]:+.3f},{vals[8]:+.3f})\n"
                    f"|X||Y||Z| = {norms[0]:.3f} {norms[1]:.3f} {norms[2]:.3f}"
                    f"  {'OK' if ok_n else '! not unit'}\n"
                    f"dots XY/YZ/ZX = {dots[0]:+.3f} {dots[1]:+.3f} {dots[2]:+.3f}"
                    f"  {'OK' if ok_o else '! not orthogonal'}\n"
                    f"det = {det:+.3f}  "
                    f"{'OK right-handed' if ok_h else '! check sign/handedness'}")
            elif self.mode == "quat" and len(vals) >= 4:
                txt = (f"mode: quaternion\n"
                       f"w={vals[0]:+.3f} x={vals[1]:+.3f} "
                       f"y={vals[2]:+.3f} z={vals[3]:+.3f}")
            elif self.mode == "vector":
                txt = (f"mode: vector\n"
                       f"x={vals[0]:+.3f} y={vals[1]:+.3f} z={vals[2]:+.3f}")
            else:
                txt = (f"mode: euler (deg)\n"
                       f"roll={vals[0]:+.2f} pitch={vals[1]:+.2f} "
                       f"yaw={vals[2]:+.2f}")
            txt += (f"\npackets: {self.reader.success_count}   "
                    f"errors: {self.reader.error_count}")
        self.label.setText(txt)

    def closeEvent(self, ev):
        self.reader.stop()
        super().closeEvent(ev)


# =============================================================================
#  Entry point
# =============================================================================

def main():
    ap = argparse.ArgumentParser(description="GPU 3D orientation visualizer")
    ap.add_argument("--port", "-p", default="COM3")
    ap.add_argument("--baudrate", "-b", type=int, default=115200)
    ap.add_argument("--mode", "-m", default="euler",
                    choices=["euler", "quat", "vector", "basis"],
                    help="how to interpret incoming numbers "
                         "(basis = 9 floats: body X,Y,Z axes in world frame)")
    ap.add_argument("--object", "-o", default="rocket",
                    choices=["rocket", "cube", "plane"])
    ap.add_argument("--stl", default=None, help="path to a binary .stl model")
    ap.add_argument("--transpose", action="store_true",
                    help="basis mode: treat the 3 vectors as rows (R^T) "
                         "instead of columns")
    ap.add_argument("--smooth", type=float, default=None,
                    help="0..1 smoothing (default 0.25; 1.0 in basis mode "
                         "so values are shown raw)")
    ap.add_argument("--list-ports", action="store_true")
    args = ap.parse_args()

    if args.list_ports:
        ports = list_ports.comports()
        if ports:
            print("Available ports:")
            for p in ports:
                print(f"  {p.device}: {p.description}")
        else:
            print("No COM ports found.")
        return

    smooth = args.smooth
    if smooth is None:
        smooth = 1.0 if args.mode == "basis" else 0.25

    reader = SerialReader(args.port, args.baudrate)
    reader.start()

    app = QtWidgets.QApplication(sys.argv)
    win = Viewer(reader, args.mode, args.object, args.stl, smooth, args.transpose)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
