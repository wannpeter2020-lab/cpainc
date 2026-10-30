#!/bin/bash
# sync_from_railway.sh
# Downloads the live Railway database and replaces the local CPAinc.sqlite.
# Usage: bash sync_from_railway.sh

set -e

RAILWAY_URL="https://cpainc.up.railway.app/admin/download-db"
LOGIN_URL="https://cpainc.up.railway.app/login"
LOCAL_DB="$(dirname "$0")/CPAinc.sqlite"
BACKUP="$(dirname "$0")/CPAinc_backup_$(date +%Y%m%d_%H%M%S).sqlite"

echo "Syncing database from Railway..."
echo ""

read -p "Username: " USERNAME
read -s -p "Password: " PASSWORD
echo ""

COOKIE_JAR=$(mktemp)
TMP_DB=$(mktemp)

# Step 1: GET login page to seed session cookie
curl -s -c "$COOKIE_JAR" "$LOGIN_URL" -o /dev/null

# Step 2: POST credentials — do NOT follow redirect (-L) so the session
#         cookie from the 302 response is captured before any redirect.
HTTP_LOGIN=$(curl -s \
  -c "$COOKIE_JAR" -b "$COOKIE_JAR" \
  -X POST "$LOGIN_URL" \
  --data-urlencode "username=$USERNAME" \
  --data-urlencode "password=$PASSWORD" \
  -w "%{http_code}" -o /dev/null)

if [ "$HTTP_LOGIN" != "302" ] && [ "$HTTP_LOGIN" != "200" ]; then
  echo "Login failed (HTTP $HTTP_LOGIN). Check your username and password."
  rm -f "$COOKIE_JAR" "$TMP_DB"
  exit 1
fi

# Step 3: Download the database using the authenticated session cookie
echo "Downloading database..."
HTTP_CODE=$(curl -s \
  -b "$COOKIE_JAR" \
  "$RAILWAY_URL" \
  -o "$TMP_DB" \
  -w "%{http_code}")

rm -f "$COOKIE_JAR"

if [ "$HTTP_CODE" != "200" ]; then
  echo "Error: server returned HTTP $HTTP_CODE"
  echo "Login may not have succeeded — try username 'peter' if you used an email."
  rm -f "$TMP_DB"
  exit 1
fi

# Verify it's a valid SQLite file
if ! file "$TMP_DB" | grep -q "SQLite"; then
  echo "Error: downloaded file is not a valid SQLite database."
  rm -f "$TMP_DB"
  exit 1
fi

# Back up current local DB
if [ -f "$LOCAL_DB" ]; then
  cp "$LOCAL_DB" "$BACKUP"
  echo "Backed up local DB to: $(basename "$BACKUP")"
fi

mv "$TMP_DB" "$LOCAL_DB"
echo ""
echo "Done! Local database is now in sync with Railway."
echo "Restart your local server to pick up the changes:"
echo "  python3 app.py"
