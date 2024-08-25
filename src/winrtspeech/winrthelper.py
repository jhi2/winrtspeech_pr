import asyncio
import logging

from win32more import FAILED, WinError
from win32more.asyncui import async_start_runner
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

    loop = async_start_runner()

    future = loop.create_future()

    task = loop.create_task(run_main_task(awaitable, future))

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

    r = asyncio.run(awaitable)

    RoUninitialize()

    return r


def start(awaitable):
    return start_mta(awaitable)


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
