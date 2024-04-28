#!/bin/bash
set -e

pwd


OO_BASE_DIR="/usr/lib64/libreoffice"

OO_IDL="$OO_BASE_DIR/sdk/bin/unoidl-write"
OO_IDL_READ="$OO_BASE_DIR/sdk/bin/unoidl-read"
OO_RDB_TYPES="$OO_BASE_DIR/program/types.rdb"
OO_RDB_API="$OO_BASE_DIR/program/types/offapi.rdb"

OO_ADDIN="MyAddin.oxt"
OO_ADDIN_REF="com.andlla.calc.MyAddin"


#
# Check Python files Compilation
#
echo check python files compilation
echo

for fn in *.py;
do
  echo compile $fn
  python -m py_compile $fn 
done

echo

#
# Interface
#

echo compile interface
echo "$OO_IDL"
echo

"$OO_IDL" "$OO_RDB_TYPES" "$OO_RDB_API" MyAddinXI.idl "MyAddinXI.rdb"

"$OO_IDL_READ" "$OO_RDB_TYPES" "$OO_RDB_API" "MyAddinXI.rdb" # dump compilation results


#
# build library
#
echo
echo package ...

rm "$OO_ADDIN" || echo no MyAddin.

zip "$OO_ADDIN" description.xml extension-description.txt MyAddin.xcu MyAddinXI.rdb META-INF/* *.py 


#
# Attempt Install
#
echo
echo "$OO_BASE_DIR/program/unopkg"

echo
echo remove ...

"$OO_BASE_DIR/program/unopkg" remove "$OO_ADDIN_REF" || echo no extension
"$OO_BASE_DIR/program/unopkg" list


echo
echo validate:
"$OO_BASE_DIR/program/unopkg" validate "$OO_ADDIN"

echo
echo install ...

"$OO_BASE_DIR/program/unopkg" add "$OO_ADDIN" || echo install failed

"$OO_BASE_DIR/program/unopkg" list

# Exception occurred: C++ code threw St9bad_alloc: std::bad_alloc at /builddir/build/BUILD/libreoffice-7.6.6.3/bridges/source/cpp_uno/gcc3_linux_x86-64/uno2cpp.cxx:243


#
# Test ODS
#

libreoffice --calc MyAddin.ods

