#! /bin/sh

clang-format $(find -name *.cpp | grep -v CompilerId | grep -v third_party) -i
clang-format $(find -name *.h | grep -v CompilerId | grep -v third_party) -i
