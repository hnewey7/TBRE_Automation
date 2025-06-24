#!/bin/bash

# Delete previous dists and builds.
rm -rf build dist *.spec

# Activate venv
PYTHON="$PWD/venv/Scripts/python.exe"

# Build all .py scripts.
for entry in "$PWD"/*
do
    if [[ $entry == *".py" ]]; then
        # Reverse entry.
        copy=${entry}
        len=${#copy}
        for((i=$len-1;i>=0;i--)); do rev="$rev${copy:$i:1}"; done

        # Cut reverse.
        reverse_cut=$(echo $rev | cut -d'/' -f 1)

        # Reverse back.
        len=${#reverse_cut}
        for((i=$len-1;i>=0;i--)); do rev_cut="$rev_cut${reverse_cut:$i:1}"; done

        # Cut again.
        script=$(echo $rev_cut | cut -d'.' -f 1)

        # Run pyinstaller.
        echo "Building $entry in $script"
        "$PYTHON" -m PyInstaller -D -w --name $script --specpath specs/ $rev_cut
        git add dist/$script

        # Make logs directory.
        logs_path=dist/${script}/logs
        mkdir -p $logs_path
        git add $logs_path

        # Copy config directory.
        cp -r config dist/$script
        git add dist/$script/config
    fi
done
