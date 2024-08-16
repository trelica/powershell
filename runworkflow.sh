#!/bin/zsh

# Function to show usage
show_usage() {
  echo "Usage: $0 -f <CsvFilePath> -s <Secret> -u <Url>"
  exit 1
}

# Parse command line arguments
while getopts ":f:s:u:" opt; do
  case $opt in
    f) csv_file="$OPTARG" ;;
    s) secret="$OPTARG" ;;
    u) url="$OPTARG" ;;
    *) show_usage ;;
  esac
done

# Check if all arguments were provided
if [ -z "$csv_file" ] || [ -z "$secret" ] || [ -z "$url" ]; then
  show_usage
fi

# Check if CSV file exists
if [ ! -f "$csv_file" ]; then
  echo "Error: CSV file not found at path: $csv_file"
  exit 1
fi

# Convert CSV to JSON
json_data=$(awk -F, '
  NR==1 {
    for (i=1; i<=NF; i++) {
      headers[i]=$i
    }
  }
  NR>1 {
    item="{"
    for (i=1; i<=NF; i++) {
      item = item "\"" headers[i] "\": \"" $i "\","
    }
    item = substr(item, 1, length(item)-1) "}"
    items = items item ","
  }
  END {
    items = substr(items, 1, length(items)-1)
    print "{ \"items\": [" items "] }"
  }
' "$csv_file")

# Emit the JSON data to be POSTed
echo "POST: "
echo "$json_data" | jq .

# POST the JSON data to the specified URL with the secret header
response=$(curl -s -X POST "$url" -H "Content-Type: application/json" -H "x-secret: $secret" -d "$json_data")

# Emit the response from the server
echo "Response:"
echo "$response" | jq .