# Diet-Recommendation-System

A Python-based GUI project developed to calculate Body Mass Index (BMI) and recommend personalized diet plans based on age, gender, height, weight, gym activity, and lifestyle

## Key Features

- Calculates BMI using user-provided data
- Categorizes into 8 BMI ranges (from underweight to super obese)
- Recommends diet based on BMI category
- Saves data to an Excel file (`bmi_data.xlsx`)
- Beautiful Tkinter GUI with blurred background support
- Auto-calculates age from date of birth
- Stores and shows health, gym, and lifestyle preferences

---

## Tech Stack

- **Language**: Python
- **Libraries**:
  - `tkinter` – GUI framework
  - `openpyxl` – Excel file creation
  - `PIL (Pillow)` – Image processing
  - `datetime`, `os` – Standard Python utilities

---

## How It Works

1. User enters name, DOB, age, gender, weight, height, gym activity, and lifestyle.
2. BMI is calculated and categorized.
3. A matching diet recommendation is displayed.
4. All data is saved into `bmi_data.xlsx`.
5. The GUI dynamically updates with personalized results.

---

## BMI Categories & Recommendations

| BMI Range                      | Category                  | Recommendation Preview  |
|--------------------------------|---------------------------|-------------------------|
| ≤ 15.0                         | Very Severely Underweight | Calorie-dense foods     |
| 15.1 – 16.0                    | Severely Underweight      | High-protein diet       |
| 16.1 – 18.4                    | Underweight               | Whole grains, dairy     |
| 18.5 – 25.0                    | Normal                    | Balanced diet           |
| 25.1 – 30.0                    | Overweight                | More fruits & veggies   |
| 30.1 – 35.0                    | Moderately Obese          | Reduce portion sizes    |
| 35.1 – 40.0                    | Severely Obese            | Avoid processed foods   |
| > 40.0                         | Super Obese               | Consult professional    |


