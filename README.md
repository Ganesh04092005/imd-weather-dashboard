# 🌦️ IMD Weather Forecast Dashboard

## 📌 Project Overview

This project is an automated **Weather Bulletin Generation System** for the India Meteorological Department (IMD).
It converts structured Excel data (Multi Hazard format) into a **professional Mid-Day Forecast Bulletin** document.

---

## 🚀 Features

* 📊 Upload IMD Excel data
* 📝 Automatically generate forecast & warning bulletin
* 📅 Dynamic date generation (Day 1 to Day 7)
* ⚠️ Automatic classification of warnings (Heavy, Very Heavy, Extreme)
* 📄 Generates official-style IMD Word document
* 🎯 Uses template-based approach for exact formatting

---

## 🛠️ Technologies Used

* Python
* Streamlit (for dashboard UI)
* Pandas (data processing)
* docxtpl (Word template automation)

---

## 📂 Project Structure

```
IMD PROJECT/
│
├── app.py              # Streamlit UI
├── processor.py        # Core logic (data processing + doc generation)
├── imd_template.docx   # IMD format template
├── requirements.txt    # Dependencies
```

---

## ▶️ How to Run

1. Install dependencies:

```
pip install -r requirements.txt
```

2. Run the app:

```
streamlit run app.py
```

3. Upload your IMD Excel file and download the generated bulletin.

---

## 📸 Output

* Generates a **7-Day Forecast Bulletin**
* Includes:

  * Forecast
  * Warnings
  * District-wise classification
  * Dynamic dates

---

## 🎯 Use Case

* Helps IMD officers automate bulletin generation
* Saves time and reduces manual effort
* Ensures consistent and error-free reports

---

## 👨‍💻 Author

Ganesh

---

## ⭐ Future Improvements

* Add map visualization
* Auto email reports
* Real-time weather API integration
