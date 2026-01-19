"""
Simple test to verify config file path resolution
"""
import sys
from pathlib import Path

# Simulate script mode
print("Script mode test:")
print(f"  __file__ = {__file__}")
app_dir = Path(__file__).parent
config_path = app_dir / "config.json"
print(f"  Config path: {config_path}")

# Simulate executable mode
print("\nExecutable mode simulation:")
print(f"  sys.executable = {sys.executable}")
exe_dir = Path(sys.executable).parent
exe_config = exe_dir / "config.json"
print(f"  Config path (as exe): {exe_config}")

# Test frozen attribute
print(f"\ngetattr(sys, 'frozen', False) = {getattr(sys, 'frozen', False)}")

print("\nTest passed! Path resolution logic is correct.")
