import asyncio
import logging

from win32more import FAILED, WinError
from win32more.asyncui import async_start_runner, async_task
from win32more.Windows.Win32.System.WinRT import (
    RO_INIT_MULTITHREADED,
    RO_INIT_SINGLETHREADED,
    RoInitialize,
    RoUninitialize,
)
from win32more.Windows.Win32.UI.WindowsAndMessaging import (
    MSG,
    DispatchMessage,
    GetMessage,
    PostQuitMessage,
    TranslateMessage,
)

logger = logging.getLogger()


def start_sta(awaitable):
    hr = RoInitialize(RO_INIT_SINGLETHREADED)
    if FAILED(hr):
        raise WinError(hr)

    async_start_runner()

    future = asyncio.get_event_loop().create_future()

    async_task(run_main_task, [awaitable, future])

    msg = MSG()
    while GetMessage(msg, 0, 0, 0) > 0:
        TranslateMessage(msg)
        DispatchMessage(msg)

    RoUninitialize()

    return future.result()


def start_mta(awaitable):
    hr = RoInitialize(RO_INIT_MULTITHREADED)
    if FAILED(hr):
        raise WinError(hr)

    async def this_task_is_created_to_release_gil_to_allow_callback(task):
        while not task.done():
            await asyncio.sleep(1)

    async def worker():
        task1 = asyncio.get_event_loop().create_task(awaitable)
        task2 = asyncio.get_event_loop().create_task(this_task_is_created_to_release_gil_to_allow_callback(task1))
        await asyncio.wait([task1, task2])
        return task1.result()

    r = asyncio.run(worker())

    RoUninitialize()

    return r


def start(awaitable):
    return start_sta(awaitable)


async def run_main_task(awaitable, future):
    try:
        r = await awaitable
        future.set_result(r)
    except SystemExit:
        future.set_result(0)
    except Exception as e:
        future.set_exception(e)
    finally:
        PostQuitMessage(0)
