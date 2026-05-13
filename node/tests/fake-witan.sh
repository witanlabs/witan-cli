#!/usr/bin/env bash
exec python3 "$(dirname "$0")/../../python/tests/fake_witan_rpc.py" "$@"
