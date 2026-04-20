# Digital Flighting Builder

A web app for building content flight plans and exporting a polished PowerPoint slide.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open http://localhost:8501.

## Usage

1. Set the slide title, date range, and number of weeks in the sidebar.
2. Add a flight: label, color, start/end dates, channels.
3. Switch to the **Generate Slide** tab and download the PPTX.

Export/import flights as JSON to share configurations between teammates.
