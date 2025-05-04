import os
import tempfile
from google.cloud import texttospeech
from google.oauth2 import service_account

# --- Configuration ---
# Use relative path to the credentials file
CREDENTIALS_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'uploads', 'sapheb-b87c6918d4ef.json')

# --- Helper Functions ---

def _get_tts_client():
    """Initializes and returns a TextToSpeechClient."""
    if not CREDENTIALS_PATH:
        raise ValueError("CREDENTIALS_PATH is not set in the script.") # Should not happen with hardcoded path
    if not os.path.exists(CREDENTIALS_PATH):
         raise FileNotFoundError(f"Credentials file not found at: {CREDENTIALS_PATH}")

    try:
        credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH)
        client = texttospeech.TextToSpeechClient(credentials=credentials)
        return client
    except Exception as e:
        raise RuntimeError(f"Failed to initialize Google Cloud TTS client: {e}")

# --- Core Functions for Server ---

def get_available_voices():
    """
    Fetches available voices from Google Cloud Text-to-Speech.

    Returns:
        list: A list of dictionaries, each containing 'id' (voice name)
              and 'name' (display name like 'English (US), Wavenet-D').
              Returns an empty list if fetching fails.
    """
    try:
        client = _get_tts_client()
        response = client.list_voices()
        voices_list = []
        for voice in response.voices:
            # Construct a user-friendly name
            lang_name = voice.language_codes[0] # e.g., "en-US"
            voice_type = voice.name.split('-')[-1] # e.g., "Wavenet" or "Standard"
            gender = texttospeech.SsmlVoiceGender(voice.ssml_gender).name.capitalize() # e.g., "Female"
            display_name = f"{lang_name}, {voice_type}-{voice.name.split('-')[-2]}, {gender}" # e.g., "en-US, Wavenet-D, Female"

            voices_list.append({
                'id': voice.name, # e.g., "en-US-Wavenet-D"
                'name': display_name
            })
        # Sort by name for better display
        voices_list.sort(key=lambda x: x['name'])
        return voices_list
    except Exception as e:
        print(f"Error fetching TTS voices: {e}")
        return [] # Return empty list on error

def generate_tts(text: str, voice_id: str, speaking_rate: float = 1.0, pitch: float = 0.0) -> str:
    """
    Synthesizes speech from text using the specified voice ID.

    Args:
        text (str): The text to synthesize.
        voice_id (str): The voice name ID (e.g., "en-US-Wavenet-D").
        speaking_rate (float): Speaking rate (0.25 to 4.0, 1.0 is default).
        pitch (float): Speaking pitch (-20.0 to 20.0, 0.0 is default).

    Returns:
        str: The absolute path to the generated temporary MP3 file.

    Raises:
        ValueError: If input parameters are invalid.
        RuntimeError: If TTS synthesis fails.
    """
    if not text:
        raise ValueError("Text cannot be empty.")
    if not voice_id:
        raise ValueError("Voice ID cannot be empty.")

    try:
        client = _get_tts_client()

        synthesis_input = texttospeech.SynthesisInput(text=text)

        # Extract language code from voice_id (e.g., "en-US" from "en-US-Wavenet-D")
        language_code = "-".join(voice_id.split('-')[:2])

        voice_params = texttospeech.VoiceSelectionParams(
            language_code=language_code, name=voice_id
        )

        audio_config = texttospeech.AudioConfig(
            audio_encoding=texttospeech.AudioEncoding.MP3,
            speaking_rate=speaking_rate,
            pitch=pitch
        )

        response = client.synthesize_speech(
            input=synthesis_input, voice=voice_params, audio_config=audio_config
        )

        # Create a temporary file to store the audio
        # Suffix is important for MIME type detection later if needed
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as temp_audio_file:
            temp_audio_file.write(response.audio_content)
            temp_file_path = temp_audio_file.name

        print(f"TTS audio content written to temporary file: {temp_file_path}")
        return temp_file_path

    except Exception as e:
        raise RuntimeError(f"Error during TTS synthesis: {e}")

# --- Example Usage (Optional - for testing the module directly) ---
if __name__ == "__main__":
    print("Testing TTS Module...")

    # --- IMPORTANT: Set your credentials path here for direct testing ---
    # os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\arons\Documents\sapheb-b87c6918d4ef.json"
    # --- OR ensure the environment variable is set before running ---

    if not os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"):
        print("Please set the GOOGLE_APPLICATION_CREDENTIALS environment variable or uncomment the line above to test.")
    else:
        print("\nFetching available voices...")
        available_voices = get_available_voices()
        if available_voices:
            print(f"Found {len(available_voices)} voices. First 5:")
            for voice in available_voices[:5]:
                print(f"  ID: {voice['id']}, Name: {voice['name']}")

            print("\nGenerating test speech ('Hello, world!') using the first available voice...")
            try:
                test_voice_id = available_voices[0]['id']
                output_file = generate_tts("Hello, world!", test_voice_id)
                print(f"Test speech generated successfully: {output_file}")
                # You might want to play the file or check its contents here
                # os.remove(output_file) # Clean up the test file
                # print(f"Cleaned up test file: {output_file}")
            except Exception as e:
                print(f"Test speech generation failed: {e}")
        else:
            print("Could not fetch voices.")
