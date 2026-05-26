"""
auto_worker.py - Standalone daily automation script.
Can be executed directly or scheduled via Windows Task Scheduler or cron.
"""
import sys
import os

# Add current directory to path to locate reports.py
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

# Mock streamlit before importing reports to prevent any unexpected StreamlitAPIException
from unittest.mock import MagicMock
sys.modules['streamlit'] = MagicMock()

run_auto_check_logic = None

try:
    # Attempt to import reports.py, which will hit the intercept and raise ImportError
    import reports
except ImportError as e:
    # Check if this is our custom UI intercept exception
    if str(e) == "Stop import for UI execution":
        # Retrieve the exported function from sys
        run_auto_check_logic = getattr(sys, "run_auto_check_logic", None)
    else:
        print(f"Error importing reports library: {e}")
        sys.exit(1)
except Exception as e:
    print(f"Unexpected error importing reports library: {e}")
    sys.exit(1)

if not run_auto_check_logic:
    print("Error: Could not retrieve run_auto_check_logic from reports.py.")
    sys.exit(1)

if __name__ == "__main__":
    print("------------------------------------------------------------")
    print("Running automated check for new emails and generating reports...")
    try:
        run_auto_check_logic(is_automated=False)
        print("Completed report auto-check, merge, and regeneration.")
        print("------------------------------------------------------------")
    except Exception as e:
        print(f"Error executing auto check pipeline: {e}")
        sys.exit(1)
