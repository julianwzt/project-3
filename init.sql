CREATE TABLE IF NOT EXISTS Jurusan (
    id_jurusan SERIAL PRIMARY KEY, 
    nama_jurusan VARCHAR(100), 
    fakultas VARCHAR(100), 
    jenjang VARCHAR(20)
    ); 

CREATE TABLE IF NOT EXISTS Mahasiswa (
    id SERIAL PRIMARY KEY, 
    nama VARCHAR(100), 
    umur INT, 
    nim VARCHAR(20) UNIQUE, 
    tgl_lahir DATE, 
    alamat TEXT, 
    id_jurusan INT REFERENCES Jurusan(id_jurusan)
);

INSERT INTO Jurusan (nama_jurusan, fakultas, jenjang) VALUES
('Teknik Informatika', 'Fakultas Informatika', 'S1'),
('Sistem Informasi', 'Fakultas Informatika', 'S1'),
('Teknik Elektro', 'Fakultas Teknik', 'S1'),
('Teknik Mesin', 'Fakultas Teknik', 'S1'),
('Desain Komunikasi Visual', 'Fakultas Industri Kreatif', 'D4'),
('Manajemen', 'Fakultas Ekonomi dan Bisnis', 'S1'),
('Akuntansi', 'Fakultas Ekonomi dan Bisnis', 'S1');

INSERT INTO Mahasiswa (nama, umur, nim, tgl_lahir, alamat, id_jurusan) VALUES
('Budi Santoso', 21, '13519001', '2005-05-10', 'Jl. Merdeka No. 10', 1);