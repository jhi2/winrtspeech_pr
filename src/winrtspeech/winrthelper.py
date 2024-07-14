import asyncio
import logging

from win32more import FAILED, WinError
from win32more.asyncui import async_start_runner, async_task
from win32more.Windows.Win32.System.WinRT import RO_INIT_SINGLETHREADED, RoInitialize, RoUninitialize
from win32more.Windows.Win32.UI.WindowsAndMessaging import (
    MSG,
    DispatchMessage,
    GetMessage,
    PostQuitMessage,
    TranslateMessage,
)

logger = logging.getLogger()


def start(awaitable):
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
