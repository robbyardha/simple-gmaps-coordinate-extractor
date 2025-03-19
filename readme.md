# Simple Google Maps Coordinates Extractor from Excel (Node.js)

Script Node.js untuk mengambil koordinat (Longitude dan Latitude) secara otomatis dari shortlink Google Maps di file Excel (.xlsx), lalu menyimpan hasilnya ke file Excel baru.

## Package Used

- xlsx - **Untuk baca & tulis file Excel**.
- node-fetch@2 - **Untuk fetch shortlink dan follow redirect**.

```bash
npm install xlsx node-fetch@2
```

- Must Contain - **Excel file with name input.xlsx**.
- Header Excel Must Contain - **Link**. - Short Link Google Maps
- Header Excel Must Contain - **Longtitude**.
- Header Excel Must Contain - **Latitude**.

## Requirements

- Must Contain - **Excel file with name input.xlsx**.
- Header Excel Must Contain - **Link**. - Short Link Google Maps
- Header Excel Must Contain - **Longtitude**.
- Header Excel Must Contain - **Latitude**.

## Feature

- [x] Baca file Excel (format .xlsx)
- [x] Resolve shortlink https://maps.app.goo.gl/... ke URL asli
- [x] Ambil Longitude dan Latitude dari link Google Maps
- [x] Update kolom Longtitude dan Latitude di Excel
- [x] Output rapi sesuai format input

## Installation Steps

1. **Clone the Repository**

   First, clone the repository to your local machine by running:

   ```bash
   git clone https://github.com/robbyardha/simple-gmaps-coordinate-extractor.git
   ```

2. **Run Installation Command**

   Navigate to the project folder and run the installation command to set up App:

   ```bash
   npm install

   ```

3. **Setting Up Your Data**

   Open input.xlsx, and prepare your data
   Important must have header (Link, Longtitude, Latitude)

4. **Run Project**

   After the setting up data complete, open your terminal with directory project and following:

   ```bash
   node index.js

   ```

5. **Check Your output file excel**

   Check output.xlsx di project directory

## License

- This Project Open Source you can contribute to develop
