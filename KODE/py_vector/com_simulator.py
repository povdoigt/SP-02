"""
=============================================================================
COM Simulator (orientation)
=============================================================================
Sends synthetic orientation data to a COM port to test the visualizer
without hardware. Pair a virtual COM port (com0com on Windows, socat on
Linux/macOS), run this on one end and the visualizer on the other.

Usage:
    python com_simulator.py --port COM4 --mode euler   --motion tumble
    python com_simulator.py --port COM4 --mode quat    --motion spin
    python com_simulator.py --port COM4 --mode vector  --motion rotate

Requirements:
    pip install pyserial numpy
=============================================================================
"""

import argparse
import time
import math
import numpy as np
import serial
from serial.tools import list_ports


def euler_sample(t, motion):
    if motion == "spin":
        return (0.0, 0.0, (t * 90.0) % 360.0)
    if motion == "tumble":
        return (t * 60.0 % 360.0, 25.0 * math.sin(t), t * 35.0 % 360.0)
    if motion == "wobble":
        return (15.0 * math.sin(t), 15.0 * math.cos(t * 0.7), 0.0)
    return (0.0, t * 45.0 % 360.0, 0.0)


def quat_sample(t, motion):
    if motion == "spin":
        axis = np.array([0.0, 0.0, 1.0]); ang = t * math.radians(90.0)
    elif motion == "tumble":
        axis = np.array([math.sin(t * 0.3), math.cos(t * 0.5), 0.5]); ang = t * math.radians(80.0)
    else:
        axis = np.array([1.0, 0.0, 0.0]); ang = t * math.radians(60.0)
    axis = axis / np.linalg.norm(axis)
    w = math.cos(ang / 2); s = math.sin(ang / 2)
    return (w, axis[0] * s, axis[1] * s, axis[2] * s)


def vector_sample(t, motion):
    if motion == "rotate":
        return (math.cos(t), math.sin(t), 0.3 * math.sin(t * 0.5))
    if motion == "tumble":
        return (math.sin(t), math.sin(t * 1.3), math.cos(t * 0.7))
    return (0.0, 1.0, 0.0)


def _quat_to_R(w, x, y, z):
    n = math.sqrt(w * w + x * x + y * y + z * z) or 1.0
    w, x, y, z = w / n, x / n, y / n, z / n
    return np.array([
        [1 - 2 * (y * y + z * z), 2 * (x * y - z * w),     2 * (x * z + y * w)],
        [2 * (x * y + z * w),     1 - 2 * (x * x + z * z), 2 * (y * z - x * w)],
        [2 * (x * z - y * w),     2 * (y * z + x * w),     1 - 2 * (x * x + y * y)],
    ])


def basis_sample(t, motion):
    """Reference path the firmware should reproduce: quaternion -> R -> 3 columns.

    Emits 9 floats = body X, Y, Z axes (the columns of R) in world frame,
    in order bx_x,bx_y,bx_z, by_x,by_y,by_z, bz_x,bz_y,bz_z.
    """
    w, x, y, z = quat_sample(t, motion)
    R = _quat_to_R(w, x, y, z)
    bx, by, bz = R[:, 0], R[:, 1], R[:, 2]
    return (bx[0], bx[1], bx[2], by[0], by[1], by[2], bz[0], bz[1], bz[2])


def main():
    ap = argparse.ArgumentParser(description="Synthetic orientation source")
    ap.add_argument("--port", "-p", default="COM4")
    ap.add_argument("--baudrate", "-b", type=int, default=115200)
    ap.add_argument("--mode", "-m", default="euler", choices=["euler", "quat", "vector", "basis"])
    ap.add_argument("--motion", default="tumble",
                    help="spin | tumble | wobble | rotate (depends on mode)")
    ap.add_argument("--rate", type=float, default=50.0, help="Hz")
    ap.add_argument("--list-ports", action="store_true")
    args = ap.parse_args()

    if args.list_ports:
        for p in list_ports.comports():
            print(f"  {p.device}: {p.description}")
        return

    try:
        ser = serial.Serial(args.port, args.baudrate, timeout=1)
    except serial.SerialException as e:
        print(f"Serial error: {e}")
        print("Tip: create a virtual COM pair first.")
        print("  Windows: com0com   |   Linux/macOS: socat -d -d pty,raw,echo=0 pty,raw,echo=0")
        print(f"Available ports: {[p.device for p in list_ports.comports()]}")
        return

    print(f"Sending {args.mode}/{args.motion} to {args.port} @ {args.rate} Hz. Ctrl+C to stop.")
    dt = 1.0 / args.rate
    t = 0.0
    try:
        while True:
            if args.mode == "quat":
                vals = quat_sample(t, args.motion)
            elif args.mode == "vector":
                vals = vector_sample(t, args.motion)
            elif args.mode == "basis":
                vals = basis_sample(t, args.motion)
            else:
                vals = euler_sample(t, args.motion)
            line = ",".join(f"{v:+.4f}" for v in vals) + "\n"
            ser.write(line.encode())
            t += dt
            time.sleep(dt)
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        if ser.is_open:
            ser.close()


if __name__ == "__main__":
    main()
