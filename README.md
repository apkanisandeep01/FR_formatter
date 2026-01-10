🌾 Farmer Data Pro: Aggregator & Deduplicator
Farmer Data Pro is a high-performance Streamlit dashboard designed to process, deduplicate, and organize farmer records. It allows users to upload multiple Excel files, clean the data based on unique identifiers, and export results as either a master file or a ZIP archive split by village.

✨ Key Features
Multi-File Upload: Combine numerous Excel workbooks into a single dataset instantly.

Smart Deduplication: Groups data by Bucket ID and Village LGD Code to ensure unique records.

High-Contrast Dashboard: A custom-styled "Processing Overview" section with high visibility for key metrics.

Top 10 Preview: Always displays the top 10 processed records for quick verification.

Dual-Format Export:

Master File: Download the entire cleaned dataset in one Excel file.

Village Split: Automatically generates individual Excel files for every village and bundles them into a ZIP folder.
🛠️ How It Works
Upload: Drag and drop your .xlsx files into the upload zone.

Process: The system automatically identifies duplicates and aggregates survey numbers.

Analyze: Check the Processing Overview cards to see how many rows were removed.

Download: Choose between the full report or the ZIP folder split by village names.

🎨 UI Fixes Applied
Visibility: Fixed the "white-on-white" text issue by applying dark-themed metric cards with high-contrast labels.

Efficiency: The preview table is restricted to the top 10 rows to maintain app performance with large datasets.
