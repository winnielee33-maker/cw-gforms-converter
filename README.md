# Coach Winnie – Forms Converter

Google Forms PDF → Microsoft Forms Quick Import Word (.docx)

## Version 2: User-supplied API Key

Each user enters their own OpenAI API Key when using the app. The key is passed only to the OpenAI client for the current Streamlit session and is not written to GitHub, a file, or a database by this app.

## Deploy on Streamlit Community Cloud

1. Upload these files to a GitHub repository.
2. Create a Streamlit app from the repository.
3. Set the main file to `app.py`.
4. Deploy. You do **not** need to put `OPENAI_API_KEY` in Streamlit Secrets for this version.
5. Optional: set `OPENAI_MODEL` as an environment variable if you want to override the default model.

## User flow

1. Enter own OpenAI API Key.
2. Upload Google Forms PDF.
3. Optionally upload an answer file.
4. Choose Form or Quiz.
5. Convert.
6. Download the generated Word file.
7. In Microsoft Forms: Quick Import → Upload from this device.

## Important

- ChatGPT subscription and OpenAI API billing are separate.
- Never commit an API key to GitHub.
- Grid / Matching questions are converted into separate choice questions.
- Image/map/chart questions are marked for manual image insertion after import.
- Answers are never guessed; they are used only when explicitly present in an uploaded answer source.
