#! /bin/sh

clang-format $(find -name *.cpp | grep -v CompilerId | grep -v third_party | grep -v build) -i
clang-format $(find -name *.h | grep -v CompilerId | grep -v third_party | grep -v build) -i
autopep8 --in-place $(find -name *.py)
