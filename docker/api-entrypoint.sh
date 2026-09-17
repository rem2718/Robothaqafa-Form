#!/bin/sh
set -eu

until pnpm --filter @workspace/db run push; do
  echo "Waiting for PostgreSQL..."
  sleep 2
done

exec pnpm --filter @workspace/api-server run start