import sys
from pathlib import Path

# Let the tests import modules from src/
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
