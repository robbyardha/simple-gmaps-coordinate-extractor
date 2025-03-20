// const XLSX = require('xlsx');
// const fetch = require('node-fetch');
// const path = require('path');

// // Fungsi untuk resolve shortlink dan ambil koordinat
// async function getCoordinates(shortUrl) {
//     try {
//         const response = await fetch(shortUrl, { redirect: 'follow' });
//         const finalUrl = response.url;
//         console.log('Resolved URL:', finalUrl);

//         const regex = /@(-?\d+\.\d+),(-?\d+\.\d+)/;
//         const match = finalUrl.match(regex);
//         if (match) {
//             return { latitude: match[1], longitude: match[2] };
//         } else {
//             console.log('Koordinat tidak ditemukan');
//             return { latitude: '', longitude: '' };
//         }
//     } catch (err) {
//         console.error('Error resolving:', shortUrl, err);
//         return { latitude: '', longitude: '' };
//     }
// }

// async function processExcel() {
//     //read file input.xlsx
//     const workbook = XLSX.readFile('input.xlsx');
//     const sheetName = workbook.SheetNames[0];
//     const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

//     //loop header
//     for (let i = 0; i < worksheet.length; i++) {
//         const row = worksheet[i];
//         const link = row['Link'];  // header contain "Link"
//         if (link) {
//             console.log(`Proses Row ${i + 1}: ${link}`);
//             const coords = await getCoordinates(link);
//             row['Longtitude'] = coords.longitude;
//             row['Latitude'] = coords.latitude;
//         }
//     }

//     //export data to output.xlsx
//     const newWorksheet = XLSX.utils.json_to_sheet(worksheet);
//     const newWorkbook = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
//     XLSX.writeFile(newWorkbook, 'output.xlsx');
//     console.log('✅ Selesai! Hasil tersimpan di output.xlsx');
// }

// processExcel();






const XLSX = require('xlsx');
const fetch = require('node-fetch');
const path = require('path');

// Fungsi untuk resolve shortlink dan ambil koordinat
async function getCoordinates(shortUrl) {
    try {
        const response = await fetch(shortUrl, { redirect: 'follow' });
        const finalUrl = response.url;
        console.log('Resolved URL:', finalUrl);

        const regex = /@(-?\d+\.\d+),(-?\d+\.\d+)/;
        const match = finalUrl.match(regex);
        if (match) {
            return { latitude: match[1], longitude: match[2] };
        } else {
            console.log('Koordinat tidak ditemukan');
            return { latitude: '', longitude: '' };
        }
    } catch (err) {
        console.error('Error resolving:', shortUrl, err);
        return { latitude: '', longitude: '' };
    }
}

async function getAddressFromCoordinates(latitude, longitude) {
    try {
        const url = `https://nominatim.openstreetmap.org/reverse?lat=${latitude}&lon=${longitude}&format=json&addressdetails=1`;
        const response = await fetch(url);
        const data = await response.json();

        if (data.address) {
            const kelurahan = data.address.village || 'Tidak ditemukan';
            const kodepos = data.address.postcode || 'Tidak ditemukan';
            return { kelurahan, kodepos };
        } else {
            console.log('Tidak dapat menemukan alamat untuk koordinat tersebut.');
            return { kelurahan: '', kodepos: '' };
        }
    } catch (err) {
        console.error('Error mendapatkan alamat:', err);
        return { kelurahan: '', kodepos: '' };
    }
}


async function processExcel() {
    //read file input.xlsx
    const workbook = XLSX.readFile('input.xlsx');
    const sheetName = workbook.SheetNames[0];
    const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

    //loop header
    for (let i = 0; i < worksheet.length; i++) {
        const row = worksheet[i];
        const link = row['Link'];  // header contain "Link"
        if (link) {
            console.log(`Proses Row ${i + 1}: ${link}`);
            const coords = await getCoordinates(link);
            const wilayah = await getAddressFromCoordinates(coords.latitude, coords.longitude)
            row['Longtitude'] = coords.longitude;
            row['Latitude'] = coords.latitude;
            row['Kelurahan/desa'] = wilayah.kelurahan;
            row['Kodepos'] = wilayah.kodepos;
        }
    }

    //export data to output.xlsx
    const newWorksheet = XLSX.utils.json_to_sheet(worksheet);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
    XLSX.writeFile(newWorkbook, 'output.xlsx');
    console.log('✅ Selesai! Hasil tersimpan di output.xlsx');
}

processExcel();

