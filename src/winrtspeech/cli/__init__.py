import argparse
import asyncio
from pathlib import Path

from win32more.Windows.Media.Core import MediaSource
from win32more.Windows.Media.Playback import MediaPlayer
from win32more.Windows.Media.SpeechSynthesis import SpeechSynthesizer
from win32more.Windows.Storage.Streams import DataReader

from winrtspeech.winrthelper import start


def find_voice(name):
    for i, voice in enumerate(SpeechSynthesizer.AllVoices):
        if voice.DisplayName == name:
            return voice
    raise RuntimeError(f"Cannot find {name}")


async def save_stream(stream, outfile):
    buf = [None] * stream.Size
    with DataReader(stream) as reader:
        await reader.LoadAsync(stream.Size)
        reader.ReadBytes(buf)
    Path(outfile).write_bytes(bytes(buf))

# NOTE: With RO_INIT_MULTITHREADED, callback (MediaEnded) can be called in other thread.
async def play_stream(stream):
    with MediaPlayer() as player:
        player.Source = MediaSource.CreateFromStream(stream, stream.ContentType)
        loop = asyncio.get_event_loop()
        future = loop.create_future()
        player.MediaEnded += lambda sender, e: loop.call_soon_threadsafe(future.set_result, 0) and None
        player.Play()
        await future


def command_list(args):
    for i, voice in enumerate(SpeechSynthesizer.AllVoices):
        print(f"[{i}]", voice.DisplayName)
        print("    Description:", voice.Description)
        print("    Gender:", voice.Gender)
        print("    Id:", voice.Id)
        print("    Language:", voice.Language)


async def command_speech(args):
    with SpeechSynthesizer() as synth:
        if args.voice is not None:
            synth.Voice = find_voice(args.voice)

        if Path(args.message).is_file():
            message = Path(args.message).read_text(encoding="utf-8")
        else:
            message = args.message

        if args.ssml:
            stream = await synth.SynthesizeSsmlToStreamAsync(message)
        else:
            stream = await synth.SynthesizeTextToStreamAsync(message)

        with stream:
            if args.out is not None:
                await save_stream(stream, args.out)
            else:
                await play_stream(stream)


async def main_async():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command")

    _parser_list = subparsers.add_parser("list", help="list installed voices")

    parser_speech = subparsers.add_parser("speech", help="speech message")
    parser_speech.add_argument("--voice", help="specify voice")
    parser_speech.add_argument("--out", help="output to file (wav)")
    parser_speech.add_argument("--ssml", action="store_true", help="read message as SSML")
    parser_speech.add_argument("message", help="text or file path")

    args = parser.parse_args()

    if args.command == "list":
        command_list(args)
    elif args.command == "speech":
        await command_speech(args)
    else:
        parser.print_help()


def main():
    return start(main_async())
