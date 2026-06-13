"""
=============================================================================
COM Port Simulator
=============================================================================
Simulates a COM port with rotating vector data for testing the visualizer.

This script creates a virtual COM port loopback (Windows: VCOM, Linux: socat).

Usage:
    python com_simulator.py --port COM3 --baudrate 115200
    
    Then in another terminal:
    python vector_visualizer.py --port COM3

Requirements:
    pip install pyserial numpy
=============================================================================
"""

import sys
import argparse
import time
import numpy as np
import serial
from serial.tools import list_ports


class VectorSimulator:
    """Generates rotating vector data"""
    
    def __init__(self):
        self.t = 0.0
    
    def get_next_vector(self, dt=0.1, mode='rotating'):
        """
        Generate next vector based on mode.
        
        Modes:
        - rotating: Vector rotating in XY plane
        - spiral: Spiral motion
        - random: Random walk
        - pendulum: Pendulum motion
        """
        self.t += dt
        
        if mode == 'rotating':
            x = 50 * np.cos(self.t)
            y = 50 * np.sin(self.t)
            z = 25 * np.sin(self.t * 0.5)
        
        elif mode == 'spiral':
            r = 50 + 25 * np.sin(self.t * 0.5)
            x = r * np.cos(self.t)
            y = r * np.sin(self.t)
            z = 30 * self.t % 60 - 30
        
        elif mode == 'random':
            x = 50 * np.sin(np.random.random() * 2 * np.pi)
            y = 50 * np.cos(np.random.random() * 2 * np.pi)
            z = 50 * (np.random.random() - 0.5) * 2
        
        elif mode == 'pendulum':
            x = 50 * np.sin(self.t)
            y = 0
            z = 50 * (1 - np.cos(self.t * 2))
        
        else:
            raise ValueError(f"Unknown mode: {mode}")
        
        return [x, y, z]


class COMSimulator(object):
    """
    Virtual COM port simulator using socat (Linux) or VCOM (Windows).
    Falls back to regular serial if virtual ports not available.
    """
    
    def __init__(self, port, baudrate, mode='rotating'):
        self.port = port
        self.baudrate = baudrate
        self.mode = mode
        self.simulator = VectorSimulator()
        self.running = True
    
    def run(self):
        """Main simulation loop"""
        try:
            ser = serial.Serial(self.port, self.baudrate, timeout=1)
            print(f"✓ Connected to {self.port} @ {self.baudrate} baud")
            print(f"✓ Sending vectors in mode: {self.mode}")
            print("\nVector data:")
            print("-" * 60)
            
            count = 0
            while self.running:
                try:
                    # Get next vector
                    vector = self.simulator.get_next_vector(mode=self.mode)
                    
                    # Format as CSV
                    data = f"{vector[0]:.4f},{vector[1]:.4f},{vector[2]:.4f}\n"
                    
                    # Send to COM port
                    ser.write(data.encode())
                    
                    # Print to console
                    count += 1
                    if count % 10 == 0:
                        mag = np.linalg.norm(vector)
                        print(f"[{count:05d}] X={vector[0]:>8.2f} Y={vector[1]:>8.2f} Z={vector[2]:>8.2f} Mag={mag:>8.2f}")
                    
                    time.sleep(0.05)  # 20 Hz
                
                except KeyboardInterrupt:
                    break
                except Exception as e:
                    print(f"⚠ Error: {e}")
        
        except serial.SerialException as e:
            print(f"✗ Serial error: {e}")
            print("\nTroubleshooting:")
            print("1. Install virtual COM port software:")
            print("   - Windows: com0com (https://sourceforge.net/projects/com0com/)")
            print("   - Linux:   sudo apt install socat")
            print("\n2. Create virtual pair:")
            print("   - Windows: COM0COM GUI → Create pair")
            print("   - Linux:   socat -d -d pty,raw,echo=0 pty,raw,echo=0")
            print(f"\n3. Run this simulator on one side of the pair")
            print(f"4. Run visualizer on the other side")
            print(f"\nAvailable ports: {[p.device for p in list_ports.comports()]}")
        
        finally:
            if 'ser' in locals() and ser.is_open:
                ser.close()
            print("\n✗ Simulator stopped")
    
    def stop(self):
        """Stop the simulator"""
        self.running = False


def main():
    parser = argparse.ArgumentParser(
        description='Simulate vector data on a COM port',
        epilog="""
Examples:
  # Simple simulator (requires virtual COM pair)
  python com_simulator.py --port COM3 --mode rotating
  
  # Different modes
  python com_simulator.py --port COM3 --mode spiral
  python com_simulator.py --port COM3 --mode pendulum
  
  # Check available ports
  python com_simulator.py --list-ports
        """
    )
    
    parser.add_argument('--port', '-p', default='COM3',
                       help='COM port to send data to (default: COM3)')
    parser.add_argument('--baudrate', '-b', type=int, default=115200,
                       help='Baud rate (default: 115200)')
    parser.add_argument('--mode', '-m', default='rotating',
                       choices=['rotating', 'spiral', 'random', 'pendulum'],
                       help='Vector motion mode (default: rotating)')
    parser.add_argument('--list-ports', action='store_true',
                       help='List available COM ports and exit')
    
    args = parser.parse_args()
    
    if args.list_ports:
        ports = list_ports.comports()
        if ports:
            print("\nAvailable COM ports:")
            for port in ports:
                print(f"  • {port.device}: {port.description}")
        else:
            print("No COM ports found!")
        return
    
    # Run simulator
    simulator = COMSimulator(args.port, args.baudrate, args.mode)
    try:
        simulator.run()
    except KeyboardInterrupt:
        print("\nShutdown...")
        simulator.stop()


if __name__ == '__main__':
    main()
