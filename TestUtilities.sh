#!/bin/bash

# a non-rubust test script to check the utility functionality

# The expectation is the two tools are inverse of each other.
# We should expect the tools to produce no differences in meaning
# and as few differences to contextual items (like localizations).

here=$(pwd)
generatorfolder="$(dirname "$(realpath $0)")"
datablocksfolder="$(dirname "$generatorfolder")/OriginalDataBlocks"

# stop on any error
set -e

cd "$datablocksfolder"
git diff
git status
# only prompt if not confirmed in terminal
if [ "$1" != "-y" ]; then
	# prompt the user to run test (since blocks will be overwritten)
	read -p "Confirm test run (blocks will be overwritten): " CONFIRM
	if [ "$CONFIRM" != "y" ] && [ "$CONFIRM" != "Y" ] && [ "$CONFIRM" != "yes" ] && [ "$CONFIRM" != "Yes" ]; then
		exit
	fi
fi

# reset blocks
git reset --hard

# run tools
cd "$generatorfolder"
echo "Start reverse utility..."
python ./LevelReverseUtility.py -v DEBUG -n "Unit 23" Evaluation Cargo Dense Vault Monster Sublimation Reckless Mother AWOL Chaos
echo "Start forward utility..."
python ./LevelUtility.py -v DEBUG -n "Unit 23.xlsx" Evaluation.xlsx Cargo.xlsx Dense.xlsx Vault.xlsx Monster.xlsx Sublimation.xlsx Reckless.xlsx Mother.xlsx AWOL.xlsx Chaos.xlsx

cd "$datablocksfolder"
echo "Diff:"
git diff
# don't reset again, leave results to be examined

cd "$here"
echo "Done."
