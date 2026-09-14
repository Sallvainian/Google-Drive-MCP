#!/usr/bin/env bash
#
# File: .claude/hooks/play-tts-elevenlabs.sh
#
# ElevenLabs TTS Provider for AgentVibes
# Uses ElevenLabs API for high-quality neural TTS
#

# Fix locale warnings
export LC_ALL=C

TEXT="$1"
VOICE_OVERRIDE="$2"  # Optional: voice ID (e.g., "wDsJlOXPqcvIUKdLXjDs")

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Default voice: Jarvis
DEFAULT_VOICE="wDsJlOXPqcvIUKdLXjDs"
DEFAULT_VOICE_NAME="Jarvis"

# Check for API key
if [[ -z "$ELEVENLABS_API_KEY" ]]; then
  echo "❌ Error: ELEVENLABS_API_KEY environment variable not set"
  echo "   Add to ~/.zshrc: export ELEVENLABS_API_KEY=\"your-api-key\""
  exit 1
fi

# Validate inputs
if [[ -z "$TEXT" ]]; then
  echo "Usage: $0 \"text to speak\" [voice_id]"
  echo ""
  echo "Default voice: $DEFAULT_VOICE_NAME ($DEFAULT_VOICE)"
  echo ""
  echo "Popular ElevenLabs voices:"
  echo "  wDsJlOXPqcvIUKdLXjDs - Jarvis (default)"
  echo "  21m00Tcm4TlvDq8ikWAM - Rachel"
  echo "  AZnzlk1XvdvUeBnXmlld - Domi"
  echo "  EXAVITQu4vr4xnSDxMaL - Bella"
  echo "  ErXwobaYiN019PkySvjV - Antoni"
  echo "  MF3mGyEYCl7XYWbV9V6O - Elli"
  echo "  TxGEqnHWrfWFTfGW9XjX - Josh"
  echo "  VR6AewLTigWG4xSOukaG - Arnold"
  echo "  pNInz6obpgDQGcFmaJgB - Adam"
  echo "  yoZ06aMxZJJ28mfd3POQ - Sam"
  exit 1
fi

# Get voice file path
get_voice_file_path() {
  local voice_file=""
  if [[ -n "$CLAUDE_PROJECT_DIR" ]] && [[ -f "$CLAUDE_PROJECT_DIR/.claude/tts-voice.txt" ]]; then
    voice_file="$CLAUDE_PROJECT_DIR/.claude/tts-voice.txt"
  elif [[ -f "$SCRIPT_DIR/../tts-voice.txt" ]]; then
    voice_file="$SCRIPT_DIR/../tts-voice.txt"
  elif [[ -f "$HOME/.claude/tts-voice.txt" ]]; then
    voice_file="$HOME/.claude/tts-voice.txt"
  fi
  echo "$voice_file"
}

# Determine voice to use
VOICE_ID=""
VOICE_NAME=""

if [[ -n "$VOICE_OVERRIDE" ]]; then
  VOICE_ID="$VOICE_OVERRIDE"
  VOICE_NAME="$VOICE_OVERRIDE"
  echo "🎤 Using voice: $VOICE_OVERRIDE (session-specific)"
else
  VOICE_FILE=$(get_voice_file_path)
  if [[ -n "$VOICE_FILE" ]] && [[ -f "$VOICE_FILE" ]]; then
    FILE_VOICE=$(cat "$VOICE_FILE" 2>/dev/null | tr -d '\n\r')
    if [[ -n "$FILE_VOICE" ]]; then
      VOICE_ID="$FILE_VOICE"
      VOICE_NAME="$FILE_VOICE"
    fi
  fi

  # Fallback to default
  if [[ -z "$VOICE_ID" ]]; then
    VOICE_ID="$DEFAULT_VOICE"
    VOICE_NAME="$DEFAULT_VOICE_NAME"
  fi
fi

# Determine audio directory
if [[ -n "$CLAUDE_PROJECT_DIR" ]]; then
  AUDIO_DIR="$CLAUDE_PROJECT_DIR/.claude/audio"
else
  CURRENT_DIR="$PWD"
  while [[ "$CURRENT_DIR" != "/" ]]; do
    if [[ -d "$CURRENT_DIR/.claude" ]]; then
      AUDIO_DIR="$CURRENT_DIR/.claude/audio"
      break
    fi
    CURRENT_DIR=$(dirname "$CURRENT_DIR")
  done
  if [[ -z "$AUDIO_DIR" ]]; then
    AUDIO_DIR="$HOME/.claude/audio"
  fi
fi

mkdir -p "$AUDIO_DIR"

# Generate unique filename
TIMESTAMP=$(date +%s)
TEMP_FILE="$AUDIO_DIR/tts-elevenlabs-${TIMESTAMP}.mp3"
FINAL_FILE="$AUDIO_DIR/tts-padded-${TIMESTAMP}.wav"

# Call ElevenLabs API
API_URL="https://api.elevenlabs.io/v1/text-to-speech/${VOICE_ID}"

HTTP_CODE=$(curl -s -w "%{http_code}" -o "$TEMP_FILE" \
  -X POST "$API_URL" \
  -H "Accept: audio/mpeg" \
  -H "Content-Type: application/json" \
  -H "xi-api-key: $ELEVENLABS_API_KEY" \
  -d "{
    \"text\": $(echo "$TEXT" | jq -Rs .),
    \"model_id\": \"eleven_monolingual_v1\",
    \"voice_settings\": {
      \"stability\": 0.5,
      \"similarity_boost\": 0.75
    }
  }")

if [[ "$HTTP_CODE" != "200" ]]; then
  ERROR_MSG=$(cat "$TEMP_FILE" 2>/dev/null)
  rm -f "$TEMP_FILE"
  echo "❌ ElevenLabs API error (HTTP $HTTP_CODE)"
  echo "   $ERROR_MSG"
  exit 2
fi

if [[ ! -f "$TEMP_FILE" ]] || [[ ! -s "$TEMP_FILE" ]]; then
  echo "❌ Failed to generate audio from ElevenLabs"
  exit 3
fi

# Convert MP3 to WAV with padding
if command -v ffmpeg &> /dev/null; then
  ffmpeg -f lavfi -i anullsrc=r=44100:cl=stereo:d=0.2 -i "$TEMP_FILE" \
    -filter_complex "[0:a][1:a]concat=n=2:v=0:a=1[out]" \
    -map "[out]" -y "$FINAL_FILE" 2>/dev/null

  if [[ -f "$FINAL_FILE" ]]; then
    rm -f "$TEMP_FILE"
    TEMP_FILE="$FINAL_FILE"
  fi
else
  # No ffmpeg - rename MP3 for output
  FINAL_FILE="$AUDIO_DIR/tts-padded-${TIMESTAMP}.mp3"
  mv "$TEMP_FILE" "$FINAL_FILE"
  TEMP_FILE="$FINAL_FILE"
fi

# Play audio
LOCK_FILE="/tmp/agentvibes-audio.lock"

# Wait for previous audio to finish (max 30 seconds)
for i in {1..60}; do
  if [ ! -f "$LOCK_FILE" ]; then
    break
  fi
  sleep 0.5
done

# Create lock and play audio
touch "$LOCK_FILE"

# Get audio duration
if command -v ffprobe &> /dev/null; then
  DURATION=$(ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$TEMP_FILE" 2>/dev/null)
  DURATION=${DURATION%.*}
else
  WORD_COUNT=$(echo "$TEXT" | wc -w)
  DURATION=$(( (WORD_COUNT * 60 / 150) + 1 ))
fi
DURATION=${DURATION:-2}

# Play audio (skip if in test mode)
if [[ "${AGENTVIBES_TEST_MODE:-false}" != "true" ]] && [[ "${AGENTVIBES_NO_PLAYBACK:-false}" != "true" ]]; then
  if [[ -n "$SSH_CONNECTION" ]] && [[ -n "$PULSE_SERVER" ]]; then
    if command -v /opt/homebrew/bin/paplay &> /dev/null; then
      /opt/homebrew/bin/paplay "$TEMP_FILE" >/dev/null 2>&1 &
      echo "🔊 Playing via PulseAudio tunnel"
    else
      afplay "$TEMP_FILE" >/dev/null 2>&1 &
    fi
  else
    afplay "$TEMP_FILE" >/dev/null 2>&1 &
  fi
fi

# Release lock after duration
(sleep $DURATION; rm -f "$LOCK_FILE") &
disown

echo "🎵 Saved to: $TEMP_FILE"
echo "🎤 Voice used: $VOICE_NAME (ElevenLabs)"
