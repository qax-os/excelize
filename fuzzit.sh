#!/bin/bash
set -xe

# Validate arguments
if [ "$#" -ne 1 ]; then
    echo "Usage: $0 <fuzz-type>"
    exit 1
fi
if [ -z "$FUZZIT_API_KEY" ]; then
    echo "Set FUZZIT_API_KEY to your Fuzzit API key"
    exit 2
fi

# Configure
NAME=excelize
TYPE=$1

# Setup
export GO111MODULE="off"
go get -u github.com/dvyukov/go-fuzz/go-fuzz github.com/dvyukov/go-fuzz/go-fuzz-build
go get -d -v -u ./...
if [ ! -f fuzzit ]; then
    wget -q -O fuzzit https://github.com/fuzzitdev/fuzzit/releases/download/v2.4.29/fuzzit_Linux_x86_64
    chmod a+x fuzzit
fi

# Fuzz
go-fuzz-build -libfuzzer -o fuzzer.a .
clang -fsanitize=fuzzer fuzzer.a -o fuzzer
./fuzzit create job --type $TYPE $NAME fuzzer
