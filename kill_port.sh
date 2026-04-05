#!/bin/bash
# Kill processes using port 5001 (or specified port)

PORT=${1:-5001}

echo "Finding processes on port $PORT..."
PIDS=$(lsof -ti tcp:$PORT 2>/dev/null)

if [ -z "$PIDS" ]; then
    echo "No processes found on port $PORT"
else
    echo "Killing PIDs: $PIDS"
    echo $PIDS | xargs kill -9 2>/dev/null
    echo "Done! Port $PORT is now free."
fi
