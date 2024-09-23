#!/bin/sh

set -e
# start cron
echo "Contenedor iniciado"
crontab /scripts/crontab.txt
/usr/sbin/crond -f -l 8
echo "crond iniciado"
