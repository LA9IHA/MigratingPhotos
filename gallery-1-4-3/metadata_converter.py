import argparse
import os
import json
from typing import Any
from weird_parser import parse

# from prettifier import FileWriter, PrettyFormatter
# Copyright (C) 2023: Mons Andrew McIvor Kvindesland

arg_parser = argparse.ArgumentParser(
    description="Crawl folder to convert metadata files recursively down folder")
arg_parser.add_argument('folder', type=str, help='folder to crawl')
arg_parser.add_argument('-n', '--name-constraint', action='store',
                        default="", help='only parse folders which contains this name constraint')
arg_parser.add_argument('-d', '--dry-run', action='store_true',
                        help='Avoid writing to any files')

args = arg_parser.parse_args()
print('folder: ' + args.folder)
folder = args.folder
is_dry_run = args.dry_run
name_constraint = args.name_constraint
failures = []
oddities = []
parse_attempts = 0


def parse_file(filename: str) -> Any:
    with open(filename, 'rb') as f:
        text = f.read().decode('cp1252', errors="replace")
    return parse(text)


def toJson(parsed: Any) -> str:
    return json.dumps(parsed, sort_keys=True, indent=2)


def write_parsed_to_file(filename: str, parsed: Any):
    with open(filename, 'w+') as writeFile:
        writeFile.write(toJson(parsed=parsed))


def handle_exception(filename, ex):
    global failures
    relativePath = "/".join(filename.split("/")[1:])
    reason = ""
    if hasattr(ex, 'message'):
        reason = ex.message
    else:
        reason = ex
    print(
        f"Could not parse filename: {relativePath}\n due to: {reason}")
    failures.append({"filename": relativePath, "reason": reason})


for root, dirs, files in os.walk(folder):
    parsedAlbumInFolder = False
    parsedPhotosInFolder = False
    albumFilename = ""
    photosFilename = ""

    if name_constraint not in root:
        continue

    if "album.dat" in files:
        albumFilename = f"{root}/album.dat"
        parse_attempts = parse_attempts + 1
        try:
            parsedAlbum = parse_file(filename=albumFilename)
            parsedAlbum["relativeFilePath"] = albumFilename
            parsedAlbumFilename = "{root}/parsed-album.json".format(
                folder=folder, root=root)
            if not is_dry_run:
                write_parsed_to_file(
                    filename=parsedAlbumFilename,
                    parsed=parsedAlbum
                )
            parsedAlbumInFolder = True
        except Exception as ex:
            handle_exception(filename=albumFilename, ex=ex)

    if "photos.dat" in files:
        photosFilename = f"{root}/photos.dat"
        parse_attempts = parse_attempts + 1
        try:
            parsedPhotos = parse_file(photosFilename)
            parsedPhotos["relativeFilePath"] = photosFilename
            parsedPhotosFilename = "{root}/parsed-photos.json".format(
                folder=folder, root=root)
            if not is_dry_run:
                write_parsed_to_file(
                    filename=parsedPhotosFilename,
                    parsed=parsedPhotos
                )
            parsedPhotosInFolder = True
        except Exception as ex:
            handle_exception(filename=photosFilename, ex=ex)

        if parsedPhotosInFolder != parsedAlbumInFolder:
            oddities.append(
                " - Found and parsed {found} but not {missing} in folder {root}".format(
                    found="photos.dat" if parsedPhotosInFolder else "album.dat",
                    missing="album.dat" if parsedPhotosInFolder else "photos.dat",
                    root=root
                )
            )


print()
print('----ODDITIES----')
for oddity in oddities:
    print(oddity)
print("There were {num_oddities} oddities".format(num_oddities=len(oddities)))
print()
print('====FAILURES====')
for failure in failures:
    print(failure['filename'] + ': ' + str(failure['reason']))
print(
    'Failures Summary: {num_failures} failures in total of {parse_attempts} parse attempts'.format(
        num_failures=len(failures),
        parse_attempts=parse_attempts)
)
