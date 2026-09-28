import React, { useState, useEffect, useRef } from 'react';
import L from 'leaflet';
import * as XLSX from 'xlsx';
import {
  AlertTriangle,
  CheckCircle2,
  Clock,
  MapPin,
  Search,
  PlusCircle,
  FileSpreadsheet,
  Filter,
  Navigation,
  Info,
  Phone,
  ShieldCheck,
  Building2,
  Camera,
  Layers,
  ChevronRight,
  RefreshCw,
  ExternalLink,
  SlidersHorizontal,
  X
} from 'lucide-react';

// Custom Map Marker Icon Fix for Leaflet in Vite
const createMarkerIcon = (color: string) => {
  return L.divIcon({
    className: 'custom-map-pin',
    html: `
      <div style="background-color: ${color}; width: 26px; height: 26px; border-radius: 50% 50% 50% 0; transform: rotate(-45deg); display: flex; align-items: center; justify-content: center; border: 2.5px solid white; box-shadow: 0 4px 8px rgba(0,0,0,0.3);">
        <div style="width: 8px; height: 8px; background-color: white; border-radius: 50%;"></div>
      </div>
    `,
    iconSize: [26, 26],
    iconAnchor: [13, 26],
    popupAnchor: [0, -26],
  });
};

interface LaporanJalan {
  id: string;
  tiket: string;
  lokasi: string;
  kecamatan: string;
  lat: number;
  lng: number;
  kategori: 'Lubang' | 'Retak' | 'Amblas' | 'Gelombang' | 'Drainase';
  keparahan: 'Ringan' | 'Sedang' | 'Berat';
  deskripsi: string;
  pelapor: string;
  telepon: string;
  tanggal: string;
  status: 'Menunggu' | 'Diverifikasi' | 'Pengerjaan' | 'Selesai';
  fotoUrl?: string;
  catatanPetugas?: string;
}

const INITIAL_REPORTS: LaporanJalan[] = [
  {
    id: '1',
    tiket: 'PLG-2026-001',
    lokasi: 'Jl. R. Soekamto No. 45 (Depan PTC Mall)',
    kecamatan: 'Ilir Timur II',
    lat: -2.9515,
    lng: 104.7645,
    kategori: 'Lubang',
    keparahan: 'Berat',
    deskripsi: 'Lubang sedalam 15 cm di lajur kiri arah simpang Patal, membahayakan pengendara motor.',
    pelapor: 'Ahmad Fauzi',
    telepon: '0812-7890-1234',
    tanggal: '2026-09-24',
    status: 'Pengerjaan',
    catatanPetugas: 'Material aspal hotmix dan tim satgas UPTD Ilir telah dijadwalkan tambal sulam.',
    fotoUrl: 'https://images.unsplash.com/photo-1515162816999-a0c47dc192f7?w=600&auto=format&fit=crop&q=80',
  },
  {
    id: '2',
    tiket: 'PLG-2026-002',
    lokasi: 'Jl. Kolonel H. Burlian KM 6.5 (Dekat Asrama Haji)',
    kecamatan: 'Sukarami',
    lat: -2.9328,
    lng: 104.7214,
    kategori: 'Gelombang',
    keparahan: 'Sedang',
    deskripsi: 'Aspal bergelombang parah akibat sering dilalui truk muatan, licin saat hujan.',
    pelapor: 'Dedi Saputra',
    telepon: '0813-6543-9876',
    tanggal: '2026-09-25',
    status: 'Diverifikasi',
    catatanPetugas: 'Telah disurvei lapangan. Masuk antrean perataan dan pelapisan ulang.',
    fotoUrl: 'https://images.unsplash.com/photo-1544620347-c4fd4a3d5957?w=600&auto=format&fit=crop&q=80',
  },
  {
    id: '3',
    tiket: 'PLG-2026-003',
    lokasi: 'Jl. Jenderal Sudirman (Simpang Charitas)',
    kecamatan: 'Ilir Timur I',
    lat: -2.9734,
    lng: 104.7578,
    kategori: 'Retak',
    keparahan: 'Ringan',
    deskripsi: 'Retak buaya di dekat lampu merah, mulai melebar akibat rembesan air hujan.',
    pelapor: 'Siti Nurhaliza',
    telepon: '0852-1122-3344',
    tanggal: '2026-09-26',
    status: 'Selesai',
    catatanPetugas: 'Selesai penambalan dan perapian oleh tim pemeliharaan rutin pada 27 Sep 2026.',
    fotoUrl: 'https://images.unsplash.com/photo-1584463699039-b9e763b03649?w=600&auto=format&fit=crop&q=80',
  },
  {
    id: '4',
    tiket: 'PLG-2026-004',
    lokasi: 'Jl. Gubernur H. Bastari (Arah Stadion Jakabaring)',
    kecamatan: 'Seberang Ulu I',
    lat: -3.0182,
    lng: 104.7891,
    kategori: 'Amblas',
    keparahan: 'Berat',
    deskripsi: 'Bahu jalan amblas sepanjang 4 meter di dekat tiang LRT setelah hujan lebat.',
    pelapor: 'Bambang Irawan',
    telepon: '0819-2233-8899',
    tanggal: '2026-09-27',
    status: 'Menunggu',
    catatanPetugas: 'Laporan baru masuk, menunggu penugasan tim verifikator lapangan.',
    fotoUrl: 'https://images.unsplash.com/photo-1621905251189-08b45d6a269e?w=600&auto=format&fit=crop&q=80',
  },
  {
    id: '5',
    tiket: 'PLG-2026-005',
    lokasi: 'Jl. Mayor Zen (Dekat Pusri Kalidoni)',
    kecamatan: 'Kalidoni',
    lat: -2.9712,
    lng: 104.7984,
    kategori: 'Drainase',
    keparahan: 'Sedang',
    deskripsi: 'Air drainase meluap mengikis aspal jalan sehingga timbul genangan dan kerikil berserakan.',
    pelapor: 'Rian Wijaya',
    telepon: '0821-4455-6677',
    tanggal: '2026-09-28',
    status: 'Pengerjaan',
    catatanPetugas: 'Koordinasi bersama tim drainase dan jalan untuk normalisasi saluran dan aspal.',
  }
];

const KECAMATAN_PALEMBANG = [
  'Semua Kecamatan',
  'Ilir Barat I',
  'Ilir Barat II',
  'Ilir Timur I',
  'Ilir Timur II',
  'Ilir Timur III',
  'Seberang Ulu I',
  'Seberang Ulu II',
  'Sukarami',
  'Sako',
  'Kalidoni',
  'Bukit Kecil',
  'Kemuning',
  'Plaju',
  'Kertapati',
  'Alang-Alang Lebar',
  'Gandus',
  'Jakabaring'
];

export default function App() {
  const [reports, setReports] = useState<LaporanJalan[]>(INITIAL_REPORTS);
  const [activeTab, setActiveTab] = useState<'map' | 'form' | 'tracking' | 'guide'>('map');
  const [selectedReport, setSelectedReport] = useState<LaporanJalan | null>(null);

  // Filters
  const [filterStatus, setFilterStatus] = useState<string>('Semua');
  const [filterKecamatan, setFilterKecamatan] = useState<string>('Semua Kecamatan');
  const [filterSeverity, setFilterSeverity] = useState<string>('Semua');
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Ticket Search
  const [searchTicketInput, setSearchTicketInput] = useState<string>('');
  const [searchTicketResult, setSearchTicketResult] = useState<LaporanJalan | null | 'not_found'>(null);

  // Form State
  const [formLokasi, setFormLokasi] = useState('');
  const [formKecamatan, setFormKecamatan] = useState('Ilir Barat I');
  const [formKategori, setFormKategori] = useState<LaporanJalan['kategori']>('Lubang');
  const [formKeparahan, setFormKeparahan] = useState<LaporanJalan['keparahan']>('Sedang');
  const [formDeskripsi, setFormDeskripsi] = useState('');
  const [formPelapor, setFormPelapor] = useState('');
  const [formTelepon, setFormTelepon] = useState('');
  const [formLat, setFormLat] = useState<number>(-2.9761);
  const [formLng, setFormLng] = useState<number>(104.7754);
  const [formSubmittedTicket, setFormSubmittedTicket] = useState<string | null>(null);

  // Leaflet Map Ref
  const mapContainerRef = useRef<HTMLDivElement | null>(null);
  const mapInstanceRef = useRef<L.Map | null>(null);
  const markersLayerRef = useRef<L.LayerGroup | null>(null);

  // Filtered Reports
  const filteredReports = reports.filter((r) => {
    if (filterStatus !== 'Semua' && r.status !== filterStatus) return false;
    if (filterKecamatan !== 'Semua Kecamatan' && r.kecamatan !== filterKecamatan) return false;
    if (filterSeverity !== 'Semua' && r.keparahan !== filterSeverity) return false;
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      const match =
        r.lokasi.toLowerCase().includes(q) ||
        r.tiket.toLowerCase().includes(q) ||
        r.kecamatan.toLowerCase().includes(q) ||
        r.deskripsi.toLowerCase().includes(q);
      if (!match) return false;
    }
    return true;
  });

  // Color mapping based on status & severity
  const getStatusBadge = (status: LaporanJalan['status']) => {
    switch (status) {
      case 'Menunggu':
        return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold bg-amber-100 text-amber-800">Menunggu</span>;
      case 'Diverifikasi':
        return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold bg-blue-100 text-blue-800">Diverifikasi</span>;
      case 'Pengerjaan':
        return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold bg-purple-100 text-purple-800">Pengerjaan</span>;
      case 'Selesai':
        return <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-semibold bg-emerald-100 text-emerald-800">Selesai</span>;
    }
  };

  const getSeverityColor = (keparahan: LaporanJalan['keparahan']) => {
    switch (keparahan) {
      case 'Ringan':
        return '#eab308'; // yellow
      case 'Sedang':
        return '#f97316'; // orange
      case 'Berat':
        return '#ef4444'; // red
    }
  };

  // Initialize Leaflet Map
  useEffect(() => {
    if (activeTab !== 'map') return;
    if (!mapContainerRef.current) return;

    if (!mapInstanceRef.current) {
      // Center on Palembang (-2.9761, 104.7754)
      const map = L.map(mapContainerRef.current, {
        center: [-2.9761, 104.7754],
        zoom: 12,
        scrollWheelZoom: true,
      });

      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
        maxZoom: 19,
      }).addTo(map);

      const markersGroup = L.layerGroup().addTo(map);
      markersLayerRef.current = markersGroup;
      mapInstanceRef.current = map;
    }

    const map = mapInstanceRef.current;
    const markers = markersLayerRef.current;

    if (map && markers) {
      markers.clearLayers();

      filteredReports.forEach((report) => {
        const marker = L.marker([report.lat, report.lng], {
          icon: createMarkerIcon(getSeverityColor(report.keparahan)),
        });

        marker.bindPopup(`
          <div style="font-family: inherit; font-size: 13px; max-width: 240px; padding: 4px;">
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px;">
              <span style="font-weight: 700; color: #1e293b;">${report.tiket}</span>
              <span style="font-size: 11px; padding: 2px 6px; border-radius: 4px; background: #f1f5f9; font-weight: 600;">${report.status}</span>
            </div>
            <p style="font-weight: 600; margin: 0 0 4px 0; color: #0f172a;">${report.lokasi}</p>
            <p style="font-size: 12px; color: #475569; margin: 0 0 6px 0;">${report.deskripsi}</p>
            <div style="font-size: 11px; color: #64748b;">
              Keparahan: <strong style="color: ${getSeverityColor(report.keparahan)}">${report.keparahan}</strong> | ${report.kecamatan}
            </div>
          </div>
        `);

        marker.on('click', () => {
          setSelectedReport(report);
        });

        markers.addLayer(marker);
      });
    }

    return () => {
      // Keep map instance cached while on map tab
    };
  }, [activeTab, filteredReports]);

  // Pan to selected report
  const focusOnReport = (report: LaporanJalan) => {
    setSelectedReport(report);
    if (activeTab !== 'map') {
      setActiveTab('map');
    }
    setTimeout(() => {
      if (mapInstanceRef.current) {
        mapInstanceRef.current.flyTo([report.lat, report.lng], 15, { duration: 1.2 });
      }
    }, 150);
  };

  // Handle Form Submission
  const handleSubmitReport = (e: React.FormEvent) => {
    e.preventDefault();
    if (!formLokasi || !formDeskripsi || !formPelapor) return;

    const newTicket = `PLG-2026-${String(reports.length + 1).padStart(3, '0')}`;
    const newReport: LaporanJalan = {
      id: String(Date.now()),
      tiket: newTicket,
      lokasi: formLokasi,
      kecamatan: formKecamatan,
      lat: formLat,
      lng: formLng,
      kategori: formKategori,
      keparahan: formKeparahan,
      deskripsi: formDeskripsi,
      pelapor: formPelapor,
      telepon: formTelepon || '-',
      tanggal: new Date().toISOString().split('T')[0],
      status: 'Menunggu',
      catatanPetugas: 'Laporan telah diterima sistem. Tim survei akan memeriksa ke lokasi.',
    };

    setReports([newReport, ...reports]);
    setFormSubmittedTicket(newTicket);
    // Reset inputs
    setFormLokasi('');
    setFormDeskripsi('');
    setFormPelapor('');
    setFormTelepon('');
  };

  // Export to Excel
  const handleExportExcel = () => {
    const dataToExport = reports.map((r, index) => ({
      No: index + 1,
      'No. Tiket': r.tiket,
      Tanggal: r.tanggal,
      'Lokasi Kerusakan': r.lokasi,
      Kecamatan: r.kecamatan,
      Kategori: r.kategori,
      Keparahan: r.keparahan,
      Status: r.status,
      'Nama Pelapor': r.pelapor,
      'Kontak Pelapor': r.telepon,
      Deskripsi: r.deskripsi,
      'Catatan Petugas': r.catatanPetugas || '-',
      Latitude: r.lat,
      Longitude: r.lng,
    }));

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Jalan Palembang');
    XLSX.writeFile(wb, `SIPELAJAR_Palembang_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  // Search Ticket
  const handleSearchTicket = (e: React.FormEvent) => {
    e.preventDefault();
    const query = searchTicketInput.trim().toUpperCase();
    if (!query) return;

    const found = reports.find((r) => r.tiket.toUpperCase() === query);
    if (found) {
      setSearchTicketResult(found);
    } else {
      setSearchTicketResult('not_found');
    }
  };

  // Stats
  const totalReports = reports.length;
  const countMenunggu = reports.filter((r) => r.status === 'Menunggu').length;
  const countPengerjaan = reports.filter((r) => r.status === 'Pengerjaan' || r.status === 'Diverifikasi').length;
  const countSelesai = reports.filter((r) => r.status === 'Selesai').length;

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col font-sans">
      {/* Top Header / App Bar */}
      <header className="bg-slate-900 text-white shadow-md border-b border-slate-800 sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-gradient-to-tr from-amber-500 to-amber-400 flex items-center justify-center shadow-lg shadow-amber-500/30">
                <Building2 className="w-5 h-5 text-slate-950 font-bold" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h1 className="font-extrabold text-base sm:text-lg tracking-tight text-white flex items-center gap-1.5">
                    SIPELAJAR <span className="text-amber-400 font-medium text-xs sm:text-sm px-1.5 py-0.5 rounded bg-amber-400/10 border border-amber-400/20">PALEMBANG</span>
                  </h1>
                </div>
                <p className="text-[11px] text-slate-400 leading-none">
                  Dinas PUPR Kota Palembang • Sistem Pelaporan Jalan Rusak
                </p>
              </div>
            </div>

            {/* Navigation Tabs */}
            <nav className="flex items-center gap-1 sm:gap-2">
              <button
                onClick={() => setActiveTab('map')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm font-medium transition-colors ${
                  activeTab === 'map'
                    ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <MapPin className="w-4 h-4" />
                <span>Peta Kerusakan</span>
              </button>
              <button
                onClick={() => setActiveTab('form')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm font-medium transition-colors ${
                  activeTab === 'form'
                    ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <PlusCircle className="w-4 h-4" />
                <span>Lapor Jalan</span>
              </button>
              <button
                onClick={() => setActiveTab('tracking')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm font-medium transition-colors ${
                  activeTab === 'tracking'
                    ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <Search className="w-4 h-4" />
                <span>Cek Tiket</span>
              </button>
              <button
                onClick={() => setActiveTab('guide')}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs sm:text-sm font-medium transition-colors ${
                  activeTab === 'guide'
                    ? 'bg-amber-500 text-slate-950 font-semibold shadow-sm'
                    : 'text-slate-300 hover:text-white hover:bg-slate-800'
                }`}
              >
                <Info className="w-4 h-4" />
                <span className="hidden md:inline">Petunjuk</span>
              </button>
            </nav>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8">
        {/* Quick Stats Bar */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4 mb-6">
          <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs flex items-center justify-between">
            <div>
              <p className="text-xs font-medium text-slate-500 uppercase tracking-wider">Total Laporan</p>
              <h3 className="text-2xl font-bold text-slate-900 mt-0.5">{totalReports}</h3>
              <p className="text-[11px] text-slate-400 mt-0.5">Seluruh Palembang</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-slate-100 flex items-center justify-center text-slate-700">
              <Layers className="w-5 h-5" />
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs flex items-center justify-between">
            <div>
              <p className="text-xs font-medium text-amber-600 uppercase tracking-wider">Menunggu</p>
              <h3 className="text-2xl font-bold text-amber-700 mt-0.5">{countMenunggu}</h3>
              <p className="text-[11px] text-slate-400 mt-0.5">Antrean verifikasi</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-amber-50 flex items-center justify-center text-amber-600">
              <Clock className="w-5 h-5" />
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs flex items-center justify-between">
            <div>
              <p className="text-xs font-medium text-blue-600 uppercase tracking-wider">Diproses</p>
              <h3 className="text-2xl font-bold text-blue-700 mt-0.5">{countPengerjaan}</h3>
              <p className="text-[11px] text-slate-400 mt-0.5">Survei & perbaikan</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-blue-50 flex items-center justify-center text-blue-600">
              <RefreshCw className="w-5 h-5" />
            </div>
          </div>

          <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs flex items-center justify-between">
            <div>
              <p className="text-xs font-medium text-emerald-600 uppercase tracking-wider">Selesai</p>
              <h3 className="text-2xl font-bold text-emerald-700 mt-0.5">{countSelesai}</h3>
              <p className="text-[11px] text-slate-400 mt-0.5">Jalan tuntas diperbaiki</p>
            </div>
            <div className="w-10 h-10 rounded-lg bg-emerald-50 flex items-center justify-center text-emerald-600">
              <CheckCircle2 className="w-5 h-5" />
            </div>
          </div>
        </div>

        {/* TAB 1: PETA & DAFTAR LAPORAN */}
        {activeTab === 'map' && (
          <div className="space-y-4">
            {/* Filter Toolbar */}
            <div className="bg-white p-4 rounded-xl border border-slate-200/80 shadow-xs">
              <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
                <div className="flex-1 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                  <div className="relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      placeholder="Cari jalan, tiket, deskripsi..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full pl-9 pr-3 py-2 text-xs sm:text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:bg-white"
                    />
                  </div>

                  <select
                    value={filterStatus}
                    onChange={(e) => setFilterStatus(e.target.value)}
                    className="px-3 py-2 text-xs sm:text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500"
                  >
                    <option value="Semua">Semua Status</option>
                    <option value="Menunggu">Menunggu</option>
                    <option value="Diverifikasi">Diverifikasi</option>
                    <option value="Pengerjaan">Pengerjaan</option>
                    <option value="Selesai">Selesai</option>
                  </select>

                  <select
                    value={filterKecamatan}
                    onChange={(e) => setFilterKecamatan(e.target.value)}
                    className="px-3 py-2 text-xs sm:text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500"
                  >
                    {KECAMATAN_PALEMBANG.map((kec) => (
                      <option key={kec} value={kec}>
                        {kec}
                      </option>
                    ))}
                  </select>

                  <select
                    value={filterSeverity}
                    onChange={(e) => setFilterSeverity(e.target.value)}
                    className="px-3 py-2 text-xs sm:text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500"
                  >
                    <option value="Semua">Semua Keparahan</option>
                    <option value="Ringan">Ringan (Kuning)</option>
                    <option value="Sedang">Sedang (Oranye)</option>
                    <option value="Berat">Berat (Merah)</option>
                  </select>
                </div>

                <div className="flex items-center gap-2 self-end lg:self-auto">
                  <button
                    onClick={handleExportExcel}
                    className="flex items-center gap-1.5 px-3 py-2 text-xs sm:text-sm font-semibold bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg shadow-xs transition-colors"
                    title="Export data ke Excel"
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                    <span>Export Excel</span>
                  </button>
                  <button
                    onClick={() => {
                      setFilterStatus('Semua');
                      setFilterKecamatan('Semua Kecamatan');
                      setFilterSeverity('Semua');
                      setSearchQuery('');
                    }}
                    className="p-2 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors border border-slate-200"
                    title="Reset Filter"
                  >
                    <RefreshCw className="w-4 h-4" />
                  </button>
                </div>
              </div>
            </div>

            {/* Split View: Map + Cards List */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
              {/* Map View */}
              <div className="lg:col-span-8 bg-white rounded-xl border border-slate-200/80 shadow-xs overflow-hidden flex flex-col min-h-[420px] lg:min-h-[580px]">
                <div className="p-3 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                  <div className="flex items-center gap-2">
                    <MapPin className="w-4 h-4 text-amber-500" />
                    <span className="text-xs font-semibold text-slate-700">
                      Peta Spasial Jalan Rusak ({filteredReports.length} Titik Ditampilkan)
                    </span>
                  </div>
                  <div className="flex items-center gap-3 text-[11px] text-slate-500">
                    <span className="flex items-center gap-1">
                      <span className="w-2.5 h-2.5 rounded-full bg-red-500 inline-block"></span> Berat
                    </span>
                    <span className="flex items-center gap-1">
                      <span className="w-2.5 h-2.5 rounded-full bg-orange-500 inline-block"></span> Sedang
                    </span>
                    <span className="flex items-center gap-1">
                      <span className="w-2.5 h-2.5 rounded-full bg-yellow-500 inline-block"></span> Ringan
                    </span>
                  </div>
                </div>
                <div ref={mapContainerRef} className="flex-1 w-full h-full min-h-[380px] lg:min-h-[520px] z-0" />
              </div>

              {/* Sidebar: Report Details & List */}
              <div className="lg:col-span-4 flex flex-col gap-4">
                {/* Active Selected Report Details if any */}
                {selectedReport && (
                  <div className="bg-white rounded-xl border-2 border-amber-400 p-4 shadow-sm relative animate-fadeIn">
                    <button
                      onClick={() => setSelectedReport(null)}
                      className="absolute top-3 right-3 text-slate-400 hover:text-slate-600 p-1"
                    >
                      <X className="w-4 h-4" />
                    </button>
                    <div className="flex items-center gap-2 mb-2">
                      <span className="font-mono text-xs font-bold text-amber-600 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                        {selectedReport.tiket}
                      </span>
                      {getStatusBadge(selectedReport.status)}
                    </div>
                    <h4 className="font-bold text-slate-900 text-sm leading-snug">{selectedReport.lokasi}</h4>
                    <p className="text-xs text-slate-500 mt-0.5">Kecamatan: {selectedReport.kecamatan}</p>

                    {selectedReport.fotoUrl && (
                      <div className="mt-2.5 rounded-lg overflow-hidden h-32 bg-slate-100 border border-slate-200">
                        <img
                          src={selectedReport.fotoUrl}
                          alt={selectedReport.lokasi}
                          className="w-full h-full object-cover"
                        />
                      </div>
                    )}

                    <div className="mt-3 text-xs bg-slate-50 p-2.5 rounded-lg border border-slate-100">
                      <p className="text-slate-700 leading-relaxed font-medium">{selectedReport.deskripsi}</p>
                    </div>

                    {selectedReport.catatanPetugas && (
                      <div className="mt-2 text-xs bg-blue-50/70 p-2.5 rounded-lg border border-blue-100 text-blue-900">
                        <strong className="block text-[11px] uppercase tracking-wider text-blue-700 mb-0.5">Catatan Tim PUPR:</strong>
                        {selectedReport.catatanPetugas}
                      </div>
                    )}

                    <div className="mt-3 pt-3 border-t border-slate-100 flex items-center justify-between text-[11px] text-slate-500">
                      <span>Pelapor: {selectedReport.pelapor}</span>
                      <span>{selectedReport.tanggal}</span>
                    </div>
                  </div>
                )}

                {/* Report Feed List */}
                <div className="bg-white rounded-xl border border-slate-200/80 shadow-xs flex flex-col flex-1 max-h-[580px]">
                  <div className="p-3 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                    <span className="text-xs font-semibold text-slate-700">Daftar Titik Laporan</span>
                    <span className="text-xs text-slate-500 font-mono">{filteredReports.length} Data</span>
                  </div>

                  <div className="divide-y divide-slate-100 overflow-y-auto flex-1 p-1">
                    {filteredReports.length === 0 ? (
                      <div className="p-8 text-center text-slate-400">
                        <AlertTriangle className="w-8 h-8 mx-auto mb-2 text-amber-500 opacity-60" />
                        <p className="text-xs">Tidak ada data yang cocok dengan kriteria filter.</p>
                      </div>
                    ) : (
                      filteredReports.map((report) => (
                        <div
                          key={report.id}
                          onClick={() => focusOnReport(report)}
                          className={`p-3 rounded-lg cursor-pointer transition-all hover:bg-slate-50 ${
                            selectedReport?.id === report.id ? 'bg-amber-50/60 border border-amber-200' : ''
                          }`}
                        >
                          <div className="flex items-start justify-between gap-2">
                            <span className="font-mono text-xs font-bold text-slate-800">{report.tiket}</span>
                            {getStatusBadge(report.status)}
                          </div>
                          <p className="font-medium text-xs text-slate-900 mt-1 line-clamp-1">{report.lokasi}</p>
                          <div className="flex items-center gap-2 mt-1.5 text-[11px] text-slate-500">
                            <span className="flex items-center gap-1">
                              <span
                                className="w-2 h-2 rounded-full"
                                style={{ backgroundColor: getSeverityColor(report.keparahan) }}
                              ></span>
                              {report.kategori} ({report.keparahan})
                            </span>
                            <span>•</span>
                            <span>{report.kecamatan}</span>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: FORM LAPOR JALAN RUSAK */}
        {activeTab === 'form' && (
          <div className="max-w-3xl mx-auto">
            <div className="bg-white rounded-2xl border border-slate-200/80 shadow-xs overflow-hidden">
              <div className="bg-gradient-to-r from-slate-900 to-slate-800 p-6 text-white">
                <div className="flex items-center gap-2.5">
                  <div className="w-9 h-9 rounded-lg bg-amber-500 text-slate-950 flex items-center justify-center font-bold">
                    <PlusCircle className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-lg font-bold">Formulir Pengaduan Jalan Rusak</h3>
                    <p className="text-xs text-slate-300">
                      Sampaikan titik kerusakan jalan di wilayah Kota Palembang agar segera ditindaklanjuti.
                    </p>
                  </div>
                </div>
              </div>

              {formSubmittedTicket ? (
                <div className="p-8 text-center">
                  <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-4">
                    <CheckCircle2 className="w-8 h-8" />
                  </div>
                  <h3 className="text-xl font-bold text-slate-900">Laporan Berhasil Terkirim!</h3>
                  <p className="text-sm text-slate-600 mt-1">
                    Terima kasih atas partisipasi Anda dalam menjaga keselamatan infrastruktur jalan Kota Palembang.
                  </p>

                  <div className="my-6 inline-block bg-slate-50 border border-slate-200 p-4 rounded-xl">
                    <p className="text-xs text-slate-500 font-medium uppercase tracking-wider">Nomor Tiket Anda</p>
                    <p className="font-mono text-2xl font-extrabold text-amber-600 mt-1">{formSubmittedTicket}</p>
                    <p className="text-[11px] text-slate-400 mt-1">Simpan nomor ini untuk mengecek progres tindak lanjut.</p>
                  </div>

                  <div className="flex items-center justify-center gap-3">
                    <button
                      onClick={() => {
                        setFormSubmittedTicket(null);
                        setActiveTab('map');
                      }}
                      className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white text-xs sm:text-sm font-semibold rounded-lg transition-colors"
                    >
                      Lihat di Peta
                    </button>
                    <button
                      onClick={() => setFormSubmittedTicket(null)}
                      className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs sm:text-sm font-semibold rounded-lg transition-colors"
                    >
                      Buat Laporan Baru
                    </button>
                  </div>
                </div>
              ) : (
                <form onSubmit={handleSubmitReport} className="p-6 space-y-5">
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Nama Lengkap Pelapor <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="text"
                        required
                        placeholder="Contoh: Rian Pratama"
                        value={formPelapor}
                        onChange={(e) => setFormPelapor(e.target.value)}
                        className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:bg-white"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">
                        Nomor WhatsApp / Telepon <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="tel"
                        required
                        placeholder="Contoh: 0812-3456-7890"
                        value={formTelepon}
                        onChange={(e) => setFormTelepon(e.target.value)}
                        className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:bg-white"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Nama Jalan & Patokan Lokasi <span className="text-red-500">*</span>
                    </label>
                    <input
                      type="text"
                      required
                      placeholder="Contoh: Jl. Demang Lebar Daun (Sebelah SPBU Demang)"
                      value={formLokasi}
                      onChange={(e) => setFormLokasi(e.target.value)}
                      className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:bg-white"
                    />
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">Kecamatan</label>
                      <select
                        value={formKecamatan}
                        onChange={(e) => setFormKecamatan(e.target.value)}
                        className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500"
                      >
                        {KECAMATAN_PALEMBANG.filter((k) => k !== 'Semua Kecamatan').map((kec) => (
                          <option key={kec} value={kec}>
                            {kec}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">Kategori Kerusakan</label>
                      <select
                        value={formKategori}
                        onChange={(e) => setFormKategori(e.target.value as any)}
                        className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500"
                      >
                        <option value="Lubang">Lubang (Pothole)</option>
                        <option value="Retak">Retak Struktur</option>
                        <option value="Amblas">Amblas / Penurunan</option>
                        <option value="Gelombang">Bergelombang / Rutting</option>
                        <option value="Drainase">Genangan / Kerusakan Saluran</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-semibold text-slate-700 mb-1">Tingkat Keparahan</label>
                      <select
                        value={formKeparahan}
                        onChange={(e) => setFormKeparahan(e.target.value as any)}
                        className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500"
                      >
                        <option value="Ringan">Ringan (Masih bisa dilalui aman)</option>
                        <option value="Sedang">Sedang (Kendaraan harus pelan)</option>
                        <option value="Berat">Berat (Sangat berbahaya & rawan laka)</option>
                      </select>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-slate-700 mb-1">
                      Deskripsi Kerusakan <span className="text-red-500">*</span>
                    </label>
                    <textarea
                      required
                      rows={3}
                      placeholder="Jelaskan ukuran perkiraan lubang, kedalaman, dan bahaya yang ditimbulkan..."
                      value={formDeskripsi}
                      onChange={(e) => setFormDeskripsi(e.target.value)}
                      className="w-full px-3 py-2 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:bg-white"
                    ></textarea>
                  </div>

                  {/* Coordinate Simulation helper */}
                  <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 flex items-center justify-between">
                    <div className="flex items-center gap-2 text-xs text-slate-600">
                      <Navigation className="w-4 h-4 text-amber-500" />
                      <span>Koordinat GPS (Otomatis Palembang Center):</span>
                      <span className="font-mono font-semibold text-slate-800">
                        {formLat.toFixed(4)}, {formLng.toFixed(4)}
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        // slightly randomize around Palembang coordinates for simulation
                        const randLat = -2.9761 + (Math.random() - 0.5) * 0.05;
                        const randLng = 104.7754 + (Math.random() - 0.5) * 0.05;
                        setFormLat(randLat);
                        setFormLng(randLng);
                      }}
                      className="text-xs text-amber-600 hover:text-amber-700 font-semibold underline"
                    >
                      Acak Titik Koordinat
                    </button>
                  </div>

                  <button
                    type="submit"
                    className="w-full py-3 bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold rounded-xl shadow-md transition-colors flex items-center justify-center gap-2"
                  >
                    <CheckCircle2 className="w-5 h-5" />
                    <span>Kirim Laporan Pengaduan</span>
                  </button>
                </form>
              )}
            </div>
          </div>
        )}

        {/* TAB 3: CEK STATUS TIKET */}
        {activeTab === 'tracking' && (
          <div className="max-w-2xl mx-auto space-y-6">
            <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs">
              <h3 className="text-base font-bold text-slate-900 mb-1">Cek Status Laporan Jalan</h3>
              <p className="text-xs text-slate-500 mb-4">
                Masukkan nomor tiket pelaporan (contoh: <span className="font-mono font-semibold text-amber-600">PLG-2026-001</span>) untuk memantau progres perbaikan oleh Dinas PUPR.
              </p>

              <form onSubmit={handleSearchTicket} className="flex gap-2">
                <div className="relative flex-1">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    required
                    placeholder="Masukkan Nomor Tiket (PLG-2026-xxx)"
                    value={searchTicketInput}
                    onChange={(e) => setSearchTicketInput(e.target.value)}
                    className="w-full pl-9 pr-3 py-2.5 text-sm bg-slate-50 border border-slate-200 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-amber-500 focus:bg-white font-mono uppercase"
                  />
                </div>
                <button
                  type="submit"
                  className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white text-sm font-semibold rounded-lg transition-colors"
                >
                  Cari
                </button>
              </form>
            </div>

            {/* Tracking Result */}
            {searchTicketResult === 'not_found' && (
              <div className="bg-white p-6 rounded-2xl border border-slate-200 text-center">
                <AlertTriangle className="w-10 h-10 text-amber-500 mx-auto mb-2" />
                <h4 className="font-bold text-slate-800 text-sm">Nomor Tiket Tidak Ditemukan</h4>
                <p className="text-xs text-slate-500 mt-1">
                  Pastikan format nomor tiket sudah sesuai, seperti contoh: PLG-2026-001.
                </p>
              </div>
            )}

            {searchTicketResult && searchTicketResult !== 'not_found' && (
              <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs space-y-5 animate-fadeIn">
                <div className="flex items-start justify-between border-b border-slate-100 pb-4">
                  <div>
                    <span className="font-mono text-xs font-bold text-amber-600 bg-amber-50 px-2 py-0.5 rounded border border-amber-200">
                      {searchTicketResult.tiket}
                    </span>
                    <h4 className="font-bold text-slate-900 text-base mt-2">{searchTicketResult.lokasi}</h4>
                    <p className="text-xs text-slate-500 mt-0.5">
                      Kecamatan {searchTicketResult.kecamatan} • Dilaporkan pada {searchTicketResult.tanggal}
                    </p>
                  </div>
                  <div>{getStatusBadge(searchTicketResult.status)}</div>
                </div>

                {/* Timeline Progress */}
                <div>
                  <h5 className="text-xs font-semibold text-slate-700 uppercase tracking-wider mb-3">
                    Alur Penanganan Laporan
                  </h5>
                  <div className="relative pl-6 space-y-4 before:content-[''] before:absolute before:left-2.5 before:top-2 before:bottom-2 before:w-0.5 before:bg-slate-200">
                    <div className="relative">
                      <div className="absolute -left-6 top-1 w-4 h-4 rounded-full bg-emerald-500 ring-4 ring-white"></div>
                      <p className="text-xs font-bold text-slate-800">1. Laporan Diterima Sistem</p>
                      <p className="text-[11px] text-slate-500">Data dan foto tersimpan di server SIPELAJAR PUPR Palembang.</p>
                    </div>

                    <div className="relative">
                      <div
                        className={`absolute -left-6 top-1 w-4 h-4 rounded-full ring-4 ring-white ${
                          searchTicketResult.status !== 'Menunggu' ? 'bg-emerald-500' : 'bg-slate-300'
                        }`}
                      ></div>
                      <p className="text-xs font-bold text-slate-800">2. Verifikasi & Survei UPTD</p>
                      <p className="text-[11px] text-slate-500">Tim teknis memeriksa tingkat kerusakan dan kebutuhan material.</p>
                    </div>

                    <div className="relative">
                      <div
                        className={`absolute -left-6 top-1 w-4 h-4 rounded-full ring-4 ring-white ${
                          searchTicketResult.status === 'Pengerjaan' || searchTicketResult.status === 'Selesai'
                            ? 'bg-emerald-500'
                            : 'bg-slate-300'
                        }`}
                      ></div>
                      <p className="text-xs font-bold text-slate-800">3. Penanganan & Tambal Sulam</p>
                      <p className="text-[11px] text-slate-500">Satgas pemeliharaan jalan turun melakukan rekondisi aspal.</p>
                    </div>

                    <div className="relative">
                      <div
                        className={`absolute -left-6 top-1 w-4 h-4 rounded-full ring-4 ring-white ${
                          searchTicketResult.status === 'Selesai' ? 'bg-emerald-500' : 'bg-slate-300'
                        }`}
                      ></div>
                      <p className="text-xs font-bold text-slate-800">4. Selesai & Berita Acara</p>
                      <p className="text-[11px] text-slate-500">Jalan kembali mulus dan aman dilalui oleh warga.</p>
                    </div>
                  </div>
                </div>

                {searchTicketResult.catatanPetugas && (
                  <div className="p-3 bg-blue-50 border border-blue-100 rounded-xl text-xs text-blue-900">
                    <span className="font-semibold block mb-0.5 text-blue-800">Keterangan Petugas:</span>
                    {searchTicketResult.catatanPetugas}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* TAB 4: PETUNJUK & CARA RUN */}
        {activeTab === 'guide' && (
          <div className="max-w-3xl mx-auto space-y-6">
            {/* Cara Run Project Guide Card */}
            <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs">
              <div className="flex items-center gap-2 mb-3">
                <div className="w-8 h-8 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center font-bold">
                  ⚡
                </div>
                <h3 className="text-base font-bold text-slate-900">Cara Menjalankan (Run) Aplikasi Ini</h3>
              </div>

              <div className="space-y-4 text-xs sm:text-sm text-slate-600">
                <div className="p-4 bg-amber-50/70 border border-amber-200/80 rounded-xl">
                  <h4 className="font-bold text-amber-900 mb-1 flex items-center gap-1.5">
                    <CheckCircle2 className="w-4 h-4 text-amber-600" />
                    1. Di Google AI Studio (Paling Mudah & Otomatis)
                  </h4>
                  <p className="text-amber-800 leading-relaxed">
                    Aplikasi ini <strong>sudah otomatis berjalan langsung</strong> di lingkungan cloud preview AI Studio! Anda cukup melihat jendela <strong>Preview</strong> di sebelah kanan atau membuka tautan Development URL aplikasi.
                  </p>
                </div>

                <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl">
                  <h4 className="font-bold text-slate-900 mb-2">2. Jika Ingin Dijalankan di Komputer / Laptop Sendiri (Localhost)</h4>
                  <p className="mb-2">Pastikan sudah menginstal <strong>Node.js</strong> di komputer Anda, lalu ikuti 3 langkah berikut:</p>

                  <ol className="space-y-2 list-decimal list-inside pl-1 text-slate-700 font-mono text-xs">
                    <li className="p-2 bg-white rounded border border-slate-200">
                      <span className="font-sans font-semibold text-slate-900">Install dependensi:</span>
                      <pre className="mt-1 bg-slate-900 text-emerald-400 p-2 rounded">npm install</pre>
                    </li>
                    <li className="p-2 bg-white rounded border border-slate-200">
                      <span className="font-sans font-semibold text-slate-900">Jalankan development server:</span>
                      <pre className="mt-1 bg-slate-900 text-emerald-400 p-2 rounded">npm run dev</pre>
                    </li>
                    <li className="p-2 bg-white rounded border border-slate-200 font-sans">
                      Buka browser Anda di: <strong className="text-amber-600 font-mono">http://localhost:3000</strong>
                    </li>
                  </ol>
                </div>
              </div>
            </div>

            {/* SOP Pelaporan */}
            <div className="bg-white p-6 rounded-2xl border border-slate-200/80 shadow-xs">
              <h3 className="text-base font-bold text-slate-900 mb-3">Tentang SIPELAJAR Kota Palembang</h3>
              <p className="text-xs sm:text-sm text-slate-600 leading-relaxed mb-4">
                SIPELAJAR (Sistem Informasi Pelaporan Jalan Rusak) adalah inisiatif keterbukaan layanan publik Dinas Pekerjaan Umum dan Penataan Ruang (PUPR) Kota Palembang. Sistem ini mempermudah masyarakat Palembang melaporkan kondisi jalan rusak secara cepat, akurat, dan transparan.
              </p>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-100">
                  <h5 className="font-bold text-slate-800 mb-1">📍 Cepat & Presisi</h5>
                  <p className="text-slate-500">Didukung pemetaan geospasial OpenStreetMap/Leaflet untuk akurasi lokasi jalan.</p>
                </div>
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-100">
                  <h5 className="font-bold text-slate-800 mb-1">🔍 Tracking Real-time</h5>
                  <p className="text-slate-500">Masyarakat dapat memantau progres perbaikan dari status laporan hingga selesai.</p>
                </div>
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-100">
                  <h5 className="font-bold text-slate-800 mb-1">📊 Rekap & Ekspor</h5>
                  <p className="text-slate-500">Laporan dapat diekspor langsung ke format Microsoft Excel (.xlsx) untuk arsip dinas.</p>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="bg-slate-900 text-slate-400 text-xs py-6 border-t border-slate-800 mt-auto">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-3">
          <p>© 2026 Dinas Pekerjaan Umum & Penataan Ruang (PUPR) Kota Palembang. Hak Cipta Dilindungi.</p>
          <div className="flex items-center gap-4 text-[11px]">
            <span>Hotline Pengaduan: (0711) 351234</span>
            <span>•</span>
            <span>Kota Palembang, Sumatera Selatan</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
