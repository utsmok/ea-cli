import asyncio

from easy_access.pipeline import _run_sync


async def _sample_coro(x: int) -> int:
    await asyncio.sleep(0)
    return x + 1


def test_run_sync_from_running_loop():
    # Run _run_sync from inside an event loop by using asyncio.get_event_loop().run_until_complete
    async def inner():
        # call _run_sync which should detect the running loop and execute the coro in a thread
        result = _run_sync(_sample_coro(3))
        assert result == 4

    asyncio.run(inner())
