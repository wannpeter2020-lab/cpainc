#!/bin/bash
# sync_from_railway.sh
# Downloads the live Railway database and replaces the local CPAinc.sqlite.
# Usage: bash sync_from_railway.sh

set -e

RAILWAY_URL="https://cpainc.up.railway.app/admin/download-db"
LOCAL_DB="$(dirname "$0")/CPAinc.sqlite"
BACKUP="$(dirname "$0")/CPAinc_backup_$(date +%Y%m%d_%H%M%S).sqlite"

echo "Syncing database from Railway..."
echo ""

# Prompt for credentials
read -p "Username: " USERNAME
read -s -p "Password: " PASSWORD
echo ""

# Download using session cookie (log in first, then download)
COOKIE_JAR=$(mktemp)

# Step 1: get login page (grab CSRF token if needed — app uses session auth)
curl -s -c "$COOKIE_JAR" "$RAILWAY_URL" -o /dev/null

# Step 2: log in
LOGIN_RESP=$(curl -s -c "$COOKIE_JAR" -b "$COOKIE_JAR" \
  -X POST "https://cpainc.up.railway.app/login" \
  -d "username=$USERNAME&password=$PASSWORD" \
  -L -w "%{http_code}" -o /dev/null)

if [ "$LOGIN_RESP" != "200" ]; then
  echo "Login may have failed (HTTP $LOGIN_RESP). Trying download anyway..."
fi

# Step 3: download the database
echo "Downloading database..."
HTTP_CODE=$(curl -s -b "$COOKIE_JAR" \
  "$RAILWAY_URL" \
  -o /tmp/cpainc_railway.sqlite \
  -w "%{http_code}")

rm -f "$COOKIE_JAR"

if [ "$HTTP_CODE" != "200" ]; then
  echo "Error: server returned HTTP $HTTP_CODE"
  rm -f /tmp/cpainc_railway.sqlite
  exit 1
fi

# Verify it's a valid SQLite file
if ! file /tmp/cpainc_railway.sqlite | grep -q "SQLite"; then
  echo "Error: downloaded file is not a valid SQLite database."
  echo "You may need to log in at https://cpainc.up.railway.app first."
  rm -f /tmp/cpainc_railway.sqlite
  exit 1
fi

# Back up the current local DB
if [ -f "$LOCAL_DB" ]; then
  cp "$LOCAL_DB" "$BACKUP"
  echo "Backed up local DB to: $(basename "$BACKUP")"
fi

# Replace local DB
mv /tmp/cpainc_railway.sqlite "$LOCAL_DB"
echo ""
echo "Done! Local database is now in sync with Railway."
echo "Restart your local server to pick up the changes:"
echo "  python3 app.py"
