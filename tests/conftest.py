import sys
import os
import asyncio
import pytest
import pytest_asyncio
from tortoise import Tortoise
import logging
from loguru import logger as _lu_logger

# Ensure repository root is on PYTHONPATH for tests
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)


@pytest_asyncio.fixture(scope="function")
async def setup_test_db():
    """Setup and teardown test database for each test function."""
    # Initialize Tortoise for tests with in-memory SQLite
    await Tortoise.init(
        db_url="sqlite://:memory:",
        modules={
            "models": ["easy_access.db.models"]
        }
    )

    # Generate the schema
    await Tortoise.generate_schemas(safe=True)

    yield

    # Clean up connections and ensure the event loop is not left with
    # pending tasks or async generators which can cause pytest to hang.
    try:
        await Tortoise.close_connections()
    finally:
        # Cancel any still-running tasks (except the current one)
        try:
            loop = asyncio.get_running_loop()
            pending = [t for t in asyncio.all_tasks(loop) if t is not asyncio.current_task()]
            if pending:
                for t in pending:
                    t.cancel()
                await asyncio.gather(*pending, return_exceptions=True)

            # Shutdown async generators (Python 3.7+)
            if hasattr(loop, 'shutdown_asyncgens'):
                await loop.shutdown_asyncgens()
        except RuntimeError:
            # Event loop already closed or not running; ignore
            pass
        # Remove loguru handlers (they may spawn background threads when enqueue=True)
        try:
            _lu_logger.remove()
        except Exception:
            pass

        # Ensure standard logging systems are shutdown to release threads
        try:
            logging.shutdown()
        except Exception:
            pass


def pytest_sessionfinish(session, exitstatus):
    """Ensure any leftover asyncio tasks, Tortoise connections and logging
    resources are cleaned up when the pytest session finishes to avoid
    the process hanging. Writes diagnostics to a file so information is
    available even when the test terminal appears hung.
    """
    import os
    diag_path = os.path.join(ROOT, "hang_diagnostics.txt")
    try:
        with open(diag_path, "a", encoding="utf-8") as diag:
            from datetime import datetime, UTC

            diag.write(f"\n=== pytest_sessionfinish at {datetime.now(UTC).isoformat()}Z ===\n")

            # Threads
            try:
                import threading

                active = threading.enumerate()
                diag.write("Active threads:\n")
                for t in active:
                    diag.write(f"- {t.name} (ident={t.ident}, daemon={t.daemon})\n")
                # Dump stacks for each thread to help identify long-running work
                try:
                    import sys
                    import traceback

                    frames = sys._current_frames()
                    diag.write("\nThread stacks:\n")
                    for t in active:
                        ident = t.ident
                        diag.write(f"--- Stack for thread {t.name} (ident={ident}) ---\n")
                        if ident is not None:
                            frame = frames.get(ident)
                            if frame is not None:
                                traceback.print_stack(frame, file=diag)
                            else:
                                diag.write("(no frame available)\n")
                        else:
                            diag.write("(thread has no ident / not started)\n")
                except Exception as e:
                    diag.write(f"Could not dump thread stacks: {e}\n")
            except Exception as e:
                diag.write(f"Could not enumerate threads: {e}\n")

            # Collect pending asyncio tasks via a temporary loop
            try:
                async def _collect_tasks():
                    return [repr(t) for t in asyncio.all_tasks() if not t.done()]

                loop = asyncio.new_event_loop()
                try:
                    tasks = loop.run_until_complete(_collect_tasks())
                    if tasks:
                        diag.write("Pending asyncio tasks:\n")
                        for t in tasks:
                            diag.write(f"- {t}\n")
                    else:
                        diag.write("No pending asyncio tasks.\n")
                finally:
                    try:
                        loop.close()
                    except Exception:
                        pass
            except Exception as e:
                diag.write(f"Could not collect asyncio tasks: {e}\n")

            # Tortoise internals and final close attempt
            try:
                from tortoise import Tortoise as _T

                try:
                    inited = getattr(_T, "_inited", None)
                    apps = getattr(_T, "apps", None)
                    connections = getattr(_T, "_connections", None)
                    diag.write(f"Tortoise._inited={inited}\n")
                    try:
                        diag.write(f"Tortoise.apps={list(apps.keys()) if apps else apps}\n")
                    except Exception:
                        diag.write(f"Tortoise.apps={apps}\n")
                    diag.write(f"Tortoise._connections={connections}\n")

                    # Try one final close using a fresh event loop
                    try:
                        loop2 = asyncio.new_event_loop()
                        try:
                            loop2.run_until_complete(_T.close_connections())
                            diag.write("Tortoise.close_connections() completed in final attempt.\n")
                        except Exception as e:
                            diag.write(f"Final Tortoise.close_connections() failed: {e}\n")
                        finally:
                            try:
                                loop2.close()
                            except Exception:
                                pass
                    except Exception as e:
                        diag.write(f"Could not create loop for final Tortoise close: {e}\n")
                except Exception as e:
                    diag.write(f"Error inspecting/closing Tortoise: {e}\n")
            except Exception as e:
                diag.write(f"Tortoise import failed or not available: {e}\n")

            # Remove loguru handlers and shutdown logging
            try:
                _lu_logger.remove()
            except Exception:
                pass
            try:
                logging.shutdown()
            except Exception:
                pass

            diag.write("=== end diagnostics ===\n")
            diag.flush()
    except Exception:
        # If writing diagnostics fails, fall back to printing to stdout
        try:
            print("Failed to write hang diagnostics to file; attempting minimal stdout dump")
            import threading

            for t in threading.enumerate():
                print(f"- {t.name} (daemon={t.daemon})")
        except Exception:
            pass
    # If any non-daemon thread remains (excluding main), force process exit to
    # prevent pytest from hanging indefinitely. We prefer a clean shutdown but
    # fall back to os._exit when third-party threads (e.g. aiosqlite worker)
    # block process termination.
    try:
        import threading, os

        non_daemon = [t for t in threading.enumerate() if not t.daemon and t is not threading.main_thread()]
        if non_daemon:
            # Give a short grace period for well-behaved background threads (aiosqlite/loguru)
            # to finish their work and exit on their own. This reduces noisy forced exits.
            try:
                import time
                time.sleep(0.15)
            except Exception:
                pass

            # Re-evaluate threads after the grace period
            non_daemon_after = [t for t in threading.enumerate() if not t.daemon and t is not threading.main_thread()]
            if not non_daemon_after:
                try:
                    with open(diag_path, "a", encoding="utf-8") as diag:
                        diag.write("\nPreviously-detected non-daemon threads exited within grace period; continuing without forcing exit.\n")
                        diag.flush()
                except Exception:
                    pass
            else:
                try:
                    with open(diag_path, "a", encoding="utf-8") as diag:
                        diag.write("\nNon-daemon threads remain at session end:\n")
                        for t in non_daemon_after:
                            diag.write(f"- {t.name} (ident={t.ident})\n")
                        diag.write(f"Forcing process exit with os._exit({exitstatus or 0}) to avoid hang.\n")
                        diag.flush()
                except Exception:
                    pass
                # Force immediate process termination
                try:
                    os._exit(exitstatus or 0)
                except Exception:
                    pass
    except Exception:
        pass
