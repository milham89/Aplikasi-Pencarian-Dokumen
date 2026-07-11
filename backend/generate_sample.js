const xlsx = require('xlsx');

// Sample data
const data = [
  { "ID Barang": "B001", "Nama Barang": "Laptop Asus ROG", "Kategori": "Elektronik", "Stok": 15, "Harga": 15000000, "Lokasi": "Rak A-2" },
  { "ID Barang": "B002", "Nama Barang": "Keyboard Mechanical Keychron", "Kategori": "Aksesoris", "Stok": 30, "Harga": 1200000, "Lokasi": "Rak A-5" },
  { "ID Barang": "B003", "Nama Barang": "Monitor LG 24 inch", "Kategori": "Elektronik", "Stok": 10, "Harga": 2500000, "Lokasi": "Rak B-1" },
  { "ID Barang": "B004", "Nama Barang": "Mouse Logitech MX Master", "Kategori": "Aksesoris", "Stok": 22, "Harga": 1500000, "Lokasi": "Rak A-5" },
  { "ID Barang": "B005", "Nama Barang": "Meja Kerja Kayu Jati", "Kategori": "Furnitur", "Stok": 5, "Harga": 3500000, "Lokasi": "Lantai 2" }
];

const worksheet = xlsx.utils.json_to_sheet(data);
const workbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(workbook, worksheet, 'Stok Barang');

xlsx.writeFile(workbook, '../sample_data.xlsx');
console.log('Sample Excel file generated successfully.');
