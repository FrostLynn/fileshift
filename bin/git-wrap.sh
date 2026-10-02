#!/bin/bash
export GIT_DIR="/home/akhdan/Projects/file-converter/.git_repo"
export GIT_WORK_TREE="/home/akhdan/Projects/file-converter"
exec git "$@"
