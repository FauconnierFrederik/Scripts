#!/bin/bash

# ------------------------------------------------------------------
# Script: get_web_download_metrics.sh
# Purpose: Measure HTTP download performance for a given URL
# Usage: Zabbix external script to collect download metrics such as
#        response time, total time, and bandwidth.
# ------------------------------------------------------------------

# ------------------------
# INSTALLATION INSTRUCTIONS
# ------------------------
# Location:
#   /usr/lib/zabbix/externalscripts/get_web_download_metrics.sh
# Permissions:
#   chmod +x /usr/lib/zabbix/externalscripts/get_web_download_metrics.sh
#   chown zabbix:zabbix /usr/lib/zabbix/externalscripts/get_web_download_metrics.sh
# Note:
#   Ensure the Zabbix host has network access to the target URL.
#   Open firewall if needed.

# ------------------------
# USAGE EXAMPLES
# ------------------------
# ./get_web_download_metrics.sh <URL> [timeout_seconds] [output_path]
# Optional parameters:
#	$1 URL = TEST DOWNLOAD LINK
#   $2 timeout_seconds = Timeout in seconds (default: 90)
#   $3 output_path = Output path for downloaded file (default: /tmp/downloaded_file.tmp)
#
# Example:
# ./get_web_download_metrics.sh "http://cachefly.cachefly.net/10mb.test" 30
#
# Output Example:
# {"successful":true,"httpStatusCode":200,"bytesReceived":10485760,"downloadBandwidthMbps":172.19,"totalDownloadTimeSeconds":0.487,"timeToFirstByteSeconds":0.079}
#
# Zabbix:
# vi /etc/zabbix/zabbix_agent2.d/web_download_metrics.conf
# UserParameter=web.download.metrics[*],/usr/lib/zabbix/externalscripts/get_web_download_metrics.sh "$1" "$2" "$3"

# ------------------------
# TEST DOWNLOAD LINKS
# ------------------------

# Small Files (~10MB) - suitable for 5/10 min intervals
#   http://ipv4.download.thinkbroadband.com/10MB.zip
#   http://speedtest.tele2.net/10MB.zip
#   http://cachefly.cachefly.net/10mb.test

# Large Files (~200MB) - suitable for 60 min intervals
#   http://ipv4.download.thinkbroadband.com/200MB.zip
#   http://speedtest.tele2.net/200MB.zip
#   http://cachefly.cachefly.net/200mb.test

# ------------------------
# ARGUMENT PARSING & DEFAULTS
# ------------------------

URI="$1"
TIMEOUT="${2:-90}"
OUTPUT_PATH="${3:-/tmp/downloaded_file.tmp}"

if [ -z "$URI" ]; then
    echo '{"error":"No URI provided. Usage: script <URI> [timeout_seconds] [output_path]"}'
    exit 1
fi

# ------------------------
# DEPENDENCY CHECKS
# ------------------------

command -v curl >/dev/null || { echo '{"error":"curl is not installed"}'; exit 1; }
command -v jq   >/dev/null || { echo '{"error":"jq is not installed"}'; exit 1; }
command -v bc   >/dev/null || { echo '{"error":"bc is not installed"}'; exit 1; }

# Ensure output directory exists
OUTPUT_DIR=$(dirname "$OUTPUT_PATH")
mkdir -p "$OUTPUT_DIR"

# ------------------------
# CURL FORMAT FOR METRICS OUTPUT
# ------------------------

# Define curl output as JSON using -w
CURL_FORMAT='{ "time_namelookup": %{time_namelookup}, "time_connect": %{time_connect}, "time_appconnect": %{time_appconnect}, "time_pretransfer": %{time_pretransfer}, "time_redirect": %{time_redirect}, "time_starttransfer": %{time_starttransfer}, "time_total": %{time_total}, "size_download": %{size_download}, "http_code": %{http_code}, "url_effective": "%{url_effective}", "http_version": "%{http_version}" }'

# ------------------------
# CURL EXECUTION
# ------------------------

CURL_OUTPUT_TMP=$(mktemp)
CURL_ERROR_TMP=$(mktemp)
trap 'rm -f "$CURL_OUTPUT_TMP" "$CURL_ERROR_TMP"' EXIT

# Build curl command based on whether we want to save the file or discard it
if [ "$OUTPUT_PATH" = "/tmp/downloaded_file.tmp" ]; then
    curl_command="curl -sS --max-time $TIMEOUT -w '$CURL_FORMAT' -o /dev/null \"$URI\""
else
    curl_command="curl -sS --max-time $TIMEOUT -w '$CURL_FORMAT' -o \"$OUTPUT_PATH\" \"$URI\""
fi

# Execute the curl command
eval "$curl_command" > "$CURL_OUTPUT_TMP" 2> "$CURL_ERROR_TMP"
CURL_EXIT_CODE=$?

# Read curl's output
CURL_JSON=$(cat "$CURL_OUTPUT_TMP")
CURL_ERROR=$(cat "$CURL_ERROR_TMP")

rm -f "$CURL_OUTPUT_TMP" "$CURL_ERROR_TMP"

# ------------------------
# ERROR HANDLING
# ------------------------

if [ $CURL_EXIT_CODE -ne 0 ]; then
    ERROR_MESSAGE="Curl command failed with exit code $CURL_EXIT_CODE."
    if [ -n "$CURL_ERROR" ]; then
        ERROR_MESSAGE+=" Error: $(echo "$CURL_ERROR" | tr -d '\n' | sed 's/"/\\"/g')"
    fi

    # Try parsing partial metrics if JSON is valid
    if jq -e . >/dev/null 2>&1 <<<"$CURL_JSON"; then
        HTTP_CODE=$(echo "$CURL_JSON" | jq -r '.http_code // 0')
        TOTAL_TIME=$(echo "$CURL_JSON" | jq -r '.time_total // 0')
        TTFB_TIME=$(echo "$CURL_JSON" | jq -r '.time_starttransfer // 0')
    else
        HTTP_CODE=0
        TOTAL_TIME=0
        TTFB_TIME=0
    fi

    echo "{\"error\":\"$ERROR_MESSAGE\",\"successful\":false,\"httpStatusCode\":$HTTP_CODE,\"bytesReceived\":0,\"downloadBandwidthMbps\":0,\"totalDownloadTimeSeconds\":$TOTAL_TIME,\"timeToFirstByteSeconds\":$TTFB_TIME}"
    exit 1
fi

# ------------------------
# METRICS PROCESSING & OUTPUT
# ------------------------

# Extract fields safely with defaults
TIME_TOTAL=$(echo "$CURL_JSON" | jq -r '.time_total // 0')
SIZE_DOWNLOAD=$(echo "$CURL_JSON" | jq -r '.size_download // 0')
HTTP_CODE=$(echo "$CURL_JSON" | jq -r '.http_code // 0')
TIME_STARTTRANSFER=$(echo "$CURL_JSON" | jq -r '.time_starttransfer // 0')  # TTFB

# Compute bandwidth in Mbps
DOWNLOAD_BANDWIDTH_MBPS=0
if [[ "$TIME_TOTAL" =~ ^[0-9.]+$ && $(echo "$TIME_TOTAL > 0" | bc -l) -eq 1 ]]; then
    DOWNLOAD_BANDWIDTH_MBPS=$(echo "($SIZE_DOWNLOAD * 8) / ($TIME_TOTAL * 1000000)" | bc -l)
fi

# Round values for output
DOWNLOAD_BANDWIDTH_MBPS_ROUNDED=$(printf "%.2f" "$DOWNLOAD_BANDWIDTH_MBPS")
TIME_TOTAL_ROUNDED=$(printf "%.3f" "$TIME_TOTAL")
TIME_STARTTRANSFER_ROUNDED=$(printf "%.3f" "$TIME_STARTTRANSFER")

# Output as JSON
echo "{\"successful\":true,\"httpStatusCode\":$HTTP_CODE,\"bytesReceived\":$SIZE_DOWNLOAD,\"downloadBandwidthMbps\":$DOWNLOAD_BANDWIDTH_MBPS_ROUNDED,\"totalDownloadTimeSeconds\":$TIME_TOTAL_ROUNDED,\"timeToFirstByteSeconds\":$TIME_STARTTRANSFER_ROUNDED}"

exit 0
