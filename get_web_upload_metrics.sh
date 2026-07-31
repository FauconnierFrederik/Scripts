#!/bin/bash

# ------------------------------------------------------------------
# Script: get_web_upload_metrics.sh
# Purpose: Upload a dummy file to a specified endpoint and measure
#          upload speed, total upload time, and HTTP status.
# ------------------------------------------------------------------

# --------------------------
# INSTALLATION INSTRUCTIONS
# --------------------------
# Place this script in:
#   /usr/lib/zabbix/externalscripts/get_web_upload_metrics.sh
# Set correct permissions:
#   chmod +x /usr/lib/zabbix/externalscripts/get_web_upload_metrics.sh
#   chown zabbix:zabbix /usr/lib/zabbix/externalscripts/get_web_upload_metrics.sh
# Note:
#   Ensure the Zabbix host has network access to the target URL.
#   Open firewall if needed.

# --------------------------
# USAGE EXAMPLE
# --------------------------
# /usr/lib/zabbix/externalscripts/get_web_upload_metrics.sh "http://speedtest.tele2.net/upload.php" 10 30
# Arguments:
#   $1 - Upload endpoint URI (required)
#   $2 - File size in MB (optional, default: 10)
#   $3 - Timeout in seconds (optional, default: 90)
# 
# Output Example:
# {"successful":true,"httpStatusCode":200,"bytesSent":10485760,"uploadBandwidthMbps":41.10,"totalUploadTimeSeconds":2.041}
#
# Zabbix:
# vi /etc/zabbix/zabbix_agent2.d/web_upload_metrics.conf
# UserParameter=web.upload.metrics[*],/usr/lib/zabbix/externalscripts/get_web_upload_metrics.sh "$1" "$2" "$3"

# ------------------------
# TEST UPLOAD LINKS
# ------------------------

# Small Files (~10MB) - suitable for 5/10 min intervals
#   http://speedtest.tele2.net/upload.php


# --------------------------
# ARGUMENTS & DEFAULTS
# --------------------------
UPLOAD_URI="$1"
UPLOAD_FILE_SIZE_MB="${2:-10}"
TIMEOUT="${3:-90}"

# --------------------------
# INPUT VALIDATION
# --------------------------

# Validate URI
if [ -z "$UPLOAD_URI" ]; then
    echo '{"error":"Missing UPLOAD_URI. Usage: script <UPLOAD_URI> [upload_file_size_MB] [timeout_seconds]"}'
    exit 1
fi

# Validate file size is a positive integer
if ! [[ "$UPLOAD_FILE_SIZE_MB" =~ ^[1-9][0-9]*$ ]]; then
    echo '{"error":"UPLOAD_FILE_SIZE_MB must be a positive integer."}'
    exit 1
fi

# Validate required tools are available
command -v curl >/dev/null || { echo '{"error":"curl is not installed"}'; exit 1; }
command -v jq   >/dev/null || { echo '{"error":"jq is not installed"}'; exit 1; }
command -v bc   >/dev/null || { echo '{"error":"bc is not installed"}'; exit 1; }
command -v dd   >/dev/null || { echo '{"error":"dd is not installed"}'; exit 1; }

# --------------------------
# CREATE DUMMY UPLOAD FILE
# --------------------------
UPLOAD_FILE_PATH=$(mktemp "/tmp/${UPLOAD_FILE_SIZE_MB}MB_XXXXXX.zip")

# Ensure cleanup of temporary files on exit
trap 'rm -f "$UPLOAD_FILE_PATH" "$CURL_OUTPUT_TMP" "$CURL_ERROR_TMP"' EXIT

# Create a dummy file using /dev/zero
if ! dd if=/dev/zero of="$UPLOAD_FILE_PATH" bs=1M count="$UPLOAD_FILE_SIZE_MB" 2>/dev/null; then
    echo '{"error":"Failed to create dummy upload file. Check permissions or disk space."}'
    exit 1
fi

# --------------------------
# DEFINE CURL OUTPUT FORMAT
# --------------------------
# - Measures time, throughput, response status
CURL_FORMAT='{
    "time_namelookup": %{time_namelookup},
    "time_connect": %{time_connect},
    "time_appconnect": %{time_appconnect},
    "time_pretransfer": %{time_pretransfer},
    "time_redirect": %{time_redirect},
    "time_starttransfer": %{time_starttransfer},
    "time_total": %{time_total},
    "size_upload": %{size_upload},
    "http_code": %{http_code},
    "url_effective": "%{url_effective}",
    "http_version": "%{http_version}",
    "speed_upload": %{speed_upload}
}'

# Initialize temporary file paths for curl output/error (they will be created by eval)
CURL_OUTPUT_TMP=$(mktemp)
CURL_ERROR_TMP=$(mktemp)

# Clean up curl output/error temp files on exit
trap 'rm -f "$CURL_OUTPUT_TMP" "$CURL_ERROR_TMP"' EXIT

# --------------------------
# EXECUTE CURL UPLOAD
# --------------------------

# Upload the file using curl (default POST, adjust if needed)
curl_command="curl -sS --max-time $TIMEOUT -X POST -T \"$UPLOAD_FILE_PATH\" -o /dev/null -w '$CURL_FORMAT' \"$UPLOAD_URI\""

# Execute the curl command
eval "$curl_command" > "$CURL_OUTPUT_TMP" 2> "$CURL_ERROR_TMP"
CURL_EXIT_CODE=$?

# Read and clean the JSON output, ensuring it's a single, valid JSON string.
CURL_JSON=$(cat "$CURL_OUTPUT_TMP")
CURL_ERROR=$(cat "$CURL_ERROR_TMP")

rm -f "$CURL_OUTPUT_TMP" "$CURL_ERROR_TMP"

# --------------------------
# HANDLE CURL FAILURE
# --------------------------
if [ $CURL_EXIT_CODE -ne 0 ]; then
    ERROR_MESSAGE="Curl upload command failed with exit code $CURL_EXIT_CODE."
    if [ -n "$CURL_ERROR" ]; then
        ERROR_MESSAGE+=" Error: $(echo "$CURL_ERROR" | tr -d '\n' | sed 's/"/\\"/g')"
    fi

    # Try parsing partial metrics if JSON is valid
    if jq -e . >/dev/null 2>&1 <<<"$CURL_JSON"; then
        HTTP_CODE=$(echo "$CURL_JSON" | jq -r '.http_code // 0')
        TOTAL_TIME=$(echo "$CURL_JSON" | jq -r '.time_total // 0')
    else
        HTTP_CODE=0
        TOTAL_TIME=0
    fi

    echo "{\"error\":\"$ERROR_MESSAGE\",\"successful\":false,\"httpStatusCode\":$HTTP_CODE,\"bytesSent\":0,\"uploadBandwidthMbps\":0,\"totalUploadTimeSeconds\":$TOTAL_TIME}"
    exit 1
fi

# --------------------------
# PARSE METRICS FROM JSON
# --------------------------
# Extract fields safely with defaults
TIME_TOTAL=$(echo "$CURL_JSON" | jq -r '.time_total // 0')
SIZE_UPLOAD=$(echo "$CURL_JSON" | jq -r '.size_upload // 0')
HTTP_CODE=$(echo "$CURL_JSON" | jq -r '.http_code // 0')

# --------------------------
# CALCULATE BANDWIDTH (Mbps)
# --------------------------
UPLOAD_BANDWIDTH_MBPS=0
if (( $(echo "$TIME_TOTAL > 0" | bc -l) )); then
    UPLOAD_BANDWIDTH_MBPS=$(echo "($SIZE_UPLOAD * 8) / ($TIME_TOTAL * 1000000)" | bc -l)
fi

# Round values for cleaner output
UPLOAD_BANDWIDTH_MBPS_ROUNDED=$(printf "%.2f" "$UPLOAD_BANDWIDTH_MBPS")
TIME_TOTAL_ROUNDED=$(printf "%.3f" "$TIME_TOTAL")

# --------------------------
# OUTPUT RESULT AS JSON
# --------------------------
echo "{\"successful\":true,\"httpStatusCode\":$HTTP_CODE,\"bytesSent\":$SIZE_UPLOAD,\"uploadBandwidthMbps\":$UPLOAD_BANDWIDTH_MBPS_ROUNDED,\"totalUploadTimeSeconds\":$TIME_TOTAL_ROUNDED}"

exit 0