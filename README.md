# Q-Extractor
Docx Question Extractor & Excel Converter

A powerful Python automation tool that extracts multiple-choice questions from .docx files and converts them into a clean, structured Excel format, ready for uploading into school portals or CBT systems.

This tool saves hours of manual work by automatically detecting question numbers, options, and answers — even when they appear in different formats.

🚀 Features
✅ Intelligent Question Extraction

Supports numbered questions (1., 2), 3 , etc.)

Reads inline options like:

a) Option

(a) Option

A. Option

Detects separate-line options

Identifies Answer: X automatically

✅ Clean Excel Generation

Each processed question becomes a complete Excel row including:

S/N

Question text

Options A–E

Correct Option (converted to Excel column letter)

Score per question

Diagram/Image placeholder

✅ Score Auto-Distribution

Automatically divides a total score (default: 60) among all questions.

✅ Batch Processing

Process all .docx files inside the /source_docs folder in one run.

✅ Organized Output
source_docs/     → original .docx files  
processed/       → backup copies after processing  
done/            → final Excel files (.xlsx)

📦 Installation
1️⃣ Clone the repository
git clone https://github.com/yourusername/docx-question-extractor.git
cd docx-question-extractor

2️⃣ Install dependencies
pip install python-docx openpyxl

▶️ How to Use
1. Put your .docx question files into:
source_docs/

2. Run the script:
python extract_and_convert.py

3. Get your Excel files from:
done/


Processed copies of the original .docx files will also be stored in:

processed/

📌 Example Input (DOCX)
1. Who created the world?
a) Moses
b) God
c) Abraham
Answer: B

📌 Example Output (Excel Row)
S/N	QUESTION	OPTION 1	OPTION 2	OPTION 3	OPTION 4	OPTION 5	CORRECT OPTION	SCORE
1	Who created the world?	Moses	God	Abraham			D	2.00
🧠 How It Works

Reads .docx paragraphs

Uses regex to detect questions, options, and answer patterns

Cleans and organizes the content

Writes everything into a structured Excel sheet

Copies original documents into /processed

Saves generated Excel into /done

🛠 Built With

Python 3

python-docx

openpyxl

regex (re)

shutil and os

🎯 Ideal For

Teachers

Computer-Based Test (CBT) creators

School administrators

Education software developers

Anyone working with large question banks

💡 Future Improvements (optional)

Support for images inside .docx

Support for essay questions

Web-based upload interface

Automated answer validation
