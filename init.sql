DROP TABLE IF EXISTS public.mahasiswa;
DROP TABLE IF EXISTS public.jurusan;

CREATE TABLE public.jurusan (
    id_jurusan SERIAL PRIMARY KEY,
    nama_jurusan VARCHAR(255) NOT NULL,
    fakultas VARCHAR(255) NOT NULL,
    jenjang VARCHAR(50) NOT NULL
);

CREATE TABLE public.mahasiswa (
    id SERIAL PRIMARY KEY,
    nama VARCHAR(255) NOT NULL,
    umur INT NOT NULL,
    nim VARCHAR(50) NOT NULL UNIQUE,
    tgl_lahir DATE NOT NULL,
    alamat TEXT,
    id_jurusan INT,
    CONSTRAINT fk_mahasiswa_jurusan
        FOREIGN KEY(id_jurusan) 
        REFERENCES public.jurusan(id_jurusan)
        ON DELETE SET NULL
);

INSERT INTO public.jurusan (nama_jurusan, fakultas, jenjang) VALUES
('Teknik Informatika', 'Fakultas Informatika', 'S1'),
('Sistem Informasi', 'Fakultas Informatika', 'S1'),
('Teknik Elektro', 'Fakultas Teknik', 'S1'),
('Teknik Mesin', 'Fakultas Teknik', 'S1'),
('Desain Komunikasi Visual', 'Fakultas Industri Kreatif', 'D4'),
('Manajemen', 'Fakultas Ekonomi dan Bisnis', 'S1'),
('Akuntansi', 'Fakultas Ekonomi dan Bisnis', 'S1');

INSERT INTO public.mahasiswa (nama, umur, nim, tgl_lahir, alamat, id_jurusan) VALUES
('Budi Santoso', 21, '13519001', '2003-05-10', 'Jl. Merdeka No. 10', 1);