"""
error_handler.py
----------------
Provides safe error handling, string conversion, and retry utilities.

Author: Syrine Bouhoula / MDA Project
Date: 2025
"""

import time


class ErrorHandler:
    """Centralized error handling and retry utilities."""

    @staticmethod
    def safe_err_str(e):
        """Safely convert any exception to a readable string."""
        try:
            return str(e)
        except Exception:
            try:
                if hasattr(e, "args") and e.args:
                    return str(e.args[0])
            except Exception:
                pass
            try:
                return repr(e)
            except Exception:
                return "<unprintable error>"

    @staticmethod
    def err_is_busy(e):
        """Detect 'device busy' Canon SDK errors."""
        s = ErrorHandler.safe_err_str(e).lower()
        return ("device_busy" in s) or ("eds_err_device_busy" in s) or ("busy" in s)

    @staticmethod
    def safe_call(func, *args, **kwargs):
        """Safely call a function and return (success, result_or_error)."""
        try:
            return True, func(*args, **kwargs)
        except Exception as e:
            return False, ErrorHandler.safe_err_str(e)

    @staticmethod
    def retry_on_busy(max_retries=5, base_delay=0.3):
        """Decorator to retry functions that may raise 'busy' errors."""

        def decorator(fn):
            def wrapper(*args, **kwargs):
                for attempt in range(max_retries):
                    try:
                        return fn(*args, **kwargs)
                    except Exception as e:
                        if ErrorHandler.err_is_busy(e):
                            delay = base_delay + (attempt * 0.2)
                            print(f"⚠️ Device busy (retry {attempt + 1}/{max_retries}) — waiting {delay:.1f}s…")
                            time.sleep(delay)
                            continue
                        print(f"❌ Error in {fn.__name__}: {ErrorHandler.safe_err_str(e)}")
                        break
                print(f"❌ Failed after {max_retries} retries: {fn.__name__}")
                return None

            return wrapper

        return decorator


