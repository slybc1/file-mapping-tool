#!/bin/sh

script_dir=$(CDPATH= cd -P "$(dirname "$0")" && pwd -P) || exit 1
jar_path=$script_dir/file-mapping-aix.jar

if [ ! -r "$jar_path" ]; then
    printf '%s\n' "Cannot read $jar_path. Place the JAR beside this script." >&2
    exit 1
fi

if [ -n "${JAVA_HOME:-}" ]; then
    java_command=$JAVA_HOME/bin/java
else
    java_command=java
fi

exec "$java_command" -jar "$jar_path" "$@"
