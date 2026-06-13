"""
=============================================================================
Vector Orientation Visualizer
=============================================================================
Reads vector coordinates from a COM port (Windows) and displays them
graphically in 3D space in real-time.

Usage:
    python vector_visualizer.py [--port COM3] [--baudrate 115200]

Requirements:
    pip install pyserial matplotlib numpy scipy
=============================================================================
"""

import sys
import argparse
import threading
import queue
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from mpl_toolkits.mplot3d import Axes3D
import serial
from serial.tools import list_ports


class VectorBuffer:
    """Thread-safe buffer for vector data"""
    def __init__(self, max_history=100):
        self.data = queue.Queue()
        self.max_history = max_history
        self.history = []
        self.lock = threading.Lock()
    
    def push(self, vector):
        """Add a vector to the buffer"""
        with self.lock:
            self.history.append(vector)
            if len(self.history) > self.max_history:
                self.history.pop(0)
    
    def get_latest(self):
        """Get the most recent vector"""
        with self.lock:
            if self.history:
                return np.array(self.history[-1]).copy()
        return None
    
    def get_history(self):
        """Get all vectors in history"""
        with self.lock:
            return [np.array(v).copy() for v in self.history]


class COMReader(threading.Thread):
    """
    Thread-safe COM port reader.
    
    Expected data format:
    - CSV: "x,y,z" or "x,y,z,label"
    - Space-separated: "x y z" or "x y z label"
    - JSON: '{"x": 1.0, "y": 2.0, "z": 3.0}'
    - Tab-separated: "x\ty\tz"
    """
    def __init__(self, port, baudrate, buffer, callback=None):
        super().__init__(daemon=True)
        self.port = port
        self.baudrate = baudrate
        self.buffer = buffer
        self.callback = callback
        self.running = True
        self.ser = None
        self.error_count = 0
        self.success_count = 0
    
    def run(self):
        """Main thread loop"""
        try:
            self.ser = serial.Serial(self.port, self.baudrate, timeout=1)
            print(f"✓ Connected to {self.port} @ {self.baudrate} baud")
            print("✓ Waiting for data...")
            
            first_error_displayed = False
            
            while self.running:
                try:
                    if self.ser.in_waiting:
                        # Read raw bytes
                        raw_bytes = self.ser.readline()
                        
                        if not raw_bytes:
                            continue
                        
                        # Try to decode
                        try:
                            line = raw_bytes.decode('utf-8').strip()
                        except UnicodeDecodeError as e:
                            self.error_count += 1
                            if not first_error_displayed:
                                print(f"⚠ Decode error: {e}")
                                print(f"  Raw bytes: {raw_bytes.hex()}")
                                first_error_displayed = True
                            continue
                        
                        if not line:
                            continue
                        
                        # Parse the line
                        try:
                            vector = self._parse_line(line)
                            vector = [75 * float(v) for v in vector]  # Scale by 75 for better visualization
                            if vector is not None:
                                self.buffer.push(vector)
                                self.success_count += 1
                                
                                if self.callback:
                                    self.callback(vector)
                                
                                # Show first successful read
                                if self.success_count == 1:
                                    print(f"✓ First packet received: {vector}")
                        
                        except ValueError as e:
                            self.error_count += 1
                            if self.error_count <= 3:  # Only print first 3 errors
                                print(f"⚠ Parse error (attempt {self.error_count}): {e}")
                                print(f"  Raw line: {repr(line)}")
                                print(f"  Hex: {line.encode().hex()}")
                
                except Exception as e:
                    self.error_count += 1
                    if self.error_count <= 3:
                        print(f"⚠ Unexpected error: {e}")
        
        except serial.SerialException as e:
            print(f"✗ Serial error: {e}")
            print(f"  Available ports: {[p.device for p in list_ports.comports()]}")
        
        finally:
            if self.ser and self.ser.is_open:
                self.ser.close()
            print(f"\n✗ COM port closed ({self.success_count} packets, {self.error_count} errors)")
    
    def _parse_line(self, line):
        """Parse various data formats with robust error handling"""
        import json
        import re
        
        # =====================================================================
        # STEP 1: Clean up the line
        # =====================================================================
        
        # Remove NULL bytes (0x00) and other non-printable characters
        line = line.replace('\0', '').replace('\x00', '')
        
        # Remove any control characters except tab and newline
        line = ''.join(c for c in line if ord(c) >= 32 or c in '\t\n')
        
        # Strip whitespace, brackets, braces
        line = line.strip().strip('[]{}')
        
        if not line:
            raise ValueError("Empty line after cleanup")
        
        # =====================================================================
        # STEP 2: Try JSON format
        # =====================================================================
        if line.startswith('{'):
            try:
                data = json.loads(line)
                return [float(data['x']), float(data['y']), float(data['z'])]
            except Exception as e:
                pass  # Try other formats
        
        # =====================================================================
        # STEP 3: Try CSV (comma-separated) — MOST COMMON
        # =====================================================================
        if ',' in line:
            parts = line.split(',')
            if len(parts) >= 3:
                try:
                    values = []
                    for p in parts[:3]:
                        # Clean and convert: strip, remove +/- prefix if needed
                        p_clean = p.strip()
                        # Python's float() handles +/-, so this should work
                        values.append(float(p_clean))
                    
                    if len(values) == 3:
                        return values
                except (ValueError, AttributeError) as e:
                    pass  # Try other formats
        
        # =====================================================================
        # STEP 4: Try tab-separated
        # =====================================================================
        if '\t' in line:
            parts = line.split('\t')
            if len(parts) >= 3:
                try:
                    values = []
                    for p in parts[:3]:
                        values.append(float(p.strip()))
                    if len(values) == 3:
                        return values
                except (ValueError, AttributeError):
                    pass
        
        # =====================================================================
        # STEP 5: Try space-separated (fallback)
        # =====================================================================
        parts = line.split()
        if len(parts) >= 3:
            try:
                values = [float(p) for p in parts[:3]]
                return values
            except ValueError:
                pass
        
        # =====================================================================
        # STEP 6: Extraction par regex (last resort)
        # =====================================================================
        # Extract all numbers (including signs and decimals)
        numbers = re.findall(r'[+-]?\d+\.?\d*', line)
        if len(numbers) >= 3:
            try:
                return [float(n) for n in numbers[:3]]
            except ValueError:
                pass
        
        # =====================================================================
        # Failed to parse
        # =====================================================================
        raise ValueError(f"Could not parse line (len={len(line)}): {repr(line)}")
    
    def stop(self):
        """Stop the reader thread"""
        self.running = False
        if self.ser and self.ser.is_open:
            self.ser.close()


class VectorVisualizer:
    """
    Real-time 3D vector visualization.
    """
    def __init__(self, port, baudrate=115200, history_size=100):
        self.port = port
        self.baudrate = baudrate
        self.buffer = VectorBuffer(max_history=history_size)
        self.reader = None
        self.fig = None
        self.ax = None
        
        # Statistics
        self.stats = {
            'reads': 0,
            'errors': 0,
            'magnitude': 0.0,
            'timestamp': None
        }
    
    def _on_vector_received(self, vector):
        """Callback when a new vector is received"""
        self.stats['reads'] += 1
        self.stats['magnitude'] = np.linalg.norm(vector)
        self.stats['timestamp'] = datetime.now()
    
    def _setup_plot(self):
        """Setup matplotlib 3D plot"""
        self.fig = plt.figure(figsize=(12, 9))
        self.fig.suptitle('Vector Orientation Visualizer', fontsize=16, fontweight='bold')
        
        # 3D plot (left)
        self.ax = self.fig.add_subplot(121, projection='3d')
        self._setup_3d_axes()
        
        # Stats panel (right)
        self.ax_stats = self.fig.add_subplot(122)
        self.ax_stats.axis('off')
        self.text_stats = self.ax_stats.text(0.05, 0.95, '', 
                                            transform=self.ax_stats.transAxes,
                                            verticalalignment='top',
                                            fontfamily='monospace',
                                            fontsize=10)
    
    def _setup_3d_axes(self):
        """Configure 3D axes"""
        self.ax.set_xlabel('X', fontweight='bold')
        self.ax.set_ylabel('Y', fontweight='bold')
        self.ax.set_zlabel('Z', fontweight='bold')
        self.ax.set_title('3D Vector Space', fontweight='bold')
        
        # Set equal aspect ratio and limits
        limit = 100
        self.ax.set_xlim([-limit, limit])
        self.ax.set_ylim([-limit, limit])
        self.ax.set_zlim([-limit, limit])
        
        # Draw coordinate frame (origin)
        self.ax.quiver(0, 0, 0, limit*0.3, 0, 0, color='red', alpha=0.3, arrow_length_ratio=0.2, label='X')
        self.ax.quiver(0, 0, 0, 0, limit*0.3, 0, color='green', alpha=0.3, arrow_length_ratio=0.2, label='Y')
        self.ax.quiver(0, 0, 0, 0, 0, limit*0.3, color='blue', alpha=0.3, arrow_length_ratio=0.2, label='Z')
        
        # Grid
        self.ax.grid(True, alpha=0.3)
        self.ax.legend(loc='upper right', fontsize=8)
    
    def _update_plot(self, frame):
        """Update animation frame"""
        self.ax.clear()
        self._setup_3d_axes()
        
        # Get latest vector
        vector = self.buffer.get_latest()
        if vector is not None:
            # Draw current vector
            self.ax.quiver(0, 0, 0, vector[0], vector[1], vector[2],
                          color='purple', arrow_length_ratio=0.15, linewidth=3,
                          label='Current Vector')
            
            # Draw vector magnitude indicator (sphere)
            mag = np.linalg.norm(vector)
            u = np.linspace(0, 2*np.pi, 15)
            v = np.linspace(0, np.pi, 10)
            x_sphere = mag * np.outer(np.cos(u), np.sin(v))
            y_sphere = mag * np.outer(np.sin(u), np.sin(v))
            z_sphere = mag * np.outer(np.ones(np.size(u)), np.cos(v))
            self.ax.plot_surface(x_sphere, y_sphere, z_sphere, alpha=0.1, color='purple')
        
        # Draw history trail
        history = self.buffer.get_history()
        if len(history) > 1:
            history_array = np.array(history)
            self.ax.plot(history_array[:, 0], history_array[:, 1], history_array[:, 2],
                        'o-', color='cyan', alpha=0.5, linewidth=1, markersize=2, label='History')
        
        # Update stats
        self._update_stats(vector)
        
        return self.ax,
    
    def _update_stats(self, vector):
        """Update statistics panel"""
        if vector is None:
            status_text = "Waiting for data..."
        else:
            x, y, z = vector
            mag = np.linalg.norm(vector)
            
            # Calculate angles (spherical coordinates)
            azimuth = np.degrees(np.arctan2(y, x))
            elevation = np.degrees(np.arcsin(z / (mag + 1e-9)))
            
            status_text = f"""
VECTOR DATA
{'='*35}
X: {x:>10.4f}
Y: {y:>10.4f}
Z: {z:>10.4f}

MAGNITUDE & ANGLES
{'='*35}
Magnitude: {mag:>10.4f}
Azimuth:   {azimuth:>10.2f}° (XY plane)
Elevation: {elevation:>10.2f}° (Z angle)

STATISTICS
{'='*35}
Packets read:  {self.stats['reads']:>10}
Parse errors:  {self.reader.error_count:>10}
Last update:   {self.stats['timestamp'].strftime('%H:%M:%S') if self.stats['timestamp'] else 'N/A':>10}

PORT INFO
{'='*35}
Port:      {self.port}
Baudrate:  {self.baudrate}
History:   {len(self.buffer.get_history())}/{self.buffer.max_history}
"""
        
        self.text_stats.set_text(status_text)
    
    def start(self):
        """Start the visualizer"""
        print("\n" + "="*50)
        print("Vector Orientation Visualizer")
        print("="*50)
        print(f"Port: {self.port}")
        print(f"Baudrate: {self.baudrate}")
        print("\nExpected data formats:")
        print("  • CSV:    x,y,z")
        print("  • Space:  x y z")
        print("  • JSON:   {\"x\": 1.0, \"y\": 2.0, \"z\": 3.0}")
        print("\nControls:")
        print("  • Drag mouse to rotate view")
        print("  • Scroll to zoom")
        print("  • Close window to exit")
        print("="*50 + "\n")
        
        # Setup plot
        self._setup_plot()
        
        # Start COM reader
        self.reader = COMReader(self.port, self.baudrate, self.buffer, self._on_vector_received)
        self.reader.start()
        
        # Start animation
        anim = FuncAnimation(self.fig, self._update_plot, interval=100, blit=False)
        
        try:
            plt.tight_layout()
            plt.show()
        except KeyboardInterrupt:
            print("\nShutdown requested...")
        finally:
            self.stop()
    
    def stop(self):
        """Stop the visualizer"""
        if self.reader:
            self.reader.stop()
            self.reader.join(timeout=2)
        plt.close('all')


def list_available_ports():
    """List all available COM ports"""
    ports = list_ports.comports()
    if not ports:
        print("No COM ports found!")
        return
    
    print("\nAvailable COM ports:")
    for port in ports:
        print(f"  • {port.device}: {port.description}")


def main():
    parser = argparse.ArgumentParser(
        description='Real-time 3D vector orientation visualizer',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python vector_visualizer.py --port COM3
  python vector_visualizer.py --port COM3 --baudrate 115200
  python vector_visualizer.py --list-ports
        """
    )
    
    parser.add_argument('--port', '-p', default='COM3',
                       help='COM port (default: COM3)')
    parser.add_argument('--baudrate', '-b', type=int, default=115200,
                       help='Baud rate (default: 115200)')
    parser.add_argument('--history', type=int, default=100,
                       help='History size (default: 100)')
    parser.add_argument('--list-ports', action='store_true',
                       help='List available COM ports and exit')
    
    args = parser.parse_args()
    
    if args.list_ports:
        list_available_ports()
        return
    
    # Create and start visualizer
    visualizer = VectorVisualizer(args.port, args.baudrate, args.history)
    visualizer.start()


if __name__ == '__main__':
    main()
