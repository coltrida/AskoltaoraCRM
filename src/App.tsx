import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { XMLParser } from 'fast-xml-parser';
import { toPng } from 'html-to-image';
import { 
  FileSpreadsheet, 
  Upload, 
  TrendingUp, 
  Users, 
  Calendar, 
  AlertCircle,
  CheckCircle2,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  Download,
  Stethoscope,
  BarChart3,
  FileCode,
  Loader2,
  Building2,
  Search,
  Copy,
  MapPin,
  PhoneOff,
  FileX,
  UserCheck,
  Map as MapIcon,
  X,
  Info,
  Store,
  Warehouse
} from 'lucide-react';
import { MapContainer, TileLayer, Marker, Popup, useMap, Tooltip } from 'react-leaflet';
import MarkerClusterGroup from 'react-leaflet-cluster';
import L from 'leaflet';
import '@geoman-io/leaflet-geoman-free';
import '@geoman-io/leaflet-geoman-free/dist/leaflet-geoman.css';

// Fix Leaflet icon issue
const defaultIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
});

const redIcon = L.icon({
  iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-red.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
});

const greenIcon = L.icon({
  iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-green.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
});

const blueIcon = L.icon({
  iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-blue.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41],
});

const largeGreenIcon = L.icon({
  iconUrl: 'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-2x-green.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [35, 57],
  iconAnchor: [17, 57],
});

// Geoman Control Component
function GeomanControl({ onSelection }: { onSelection: (layer: any) => void }) {
  const map = useMap();

  useEffect(() => {
    if (!map) return;

    map.pm.addControls({
      position: 'topleft',
      drawCircleMarker: false,
      drawPolyline: false,
      drawRectangle: true,
      drawPolygon: true,
      drawCircle: true,
      drawMarker: false,
      drawText: false,
      cutPolygon: false,
      editMode: false,
      dragMode: false,
      removalMode: true,
    });

    map.on('pm:create', (e) => {
      onSelection(e.layer);
    });

    return () => {
      map.off('pm:create');
    };
  }, [map, onSelection]);

  return null;
}

// Point in Polygon Utility
function isPointInLayer(lat: number, lng: number, layer: any) {
  if (layer instanceof L.Circle) {
    const distance = layer.getLatLng().distanceTo(L.latLng(lat, lng));
    return distance <= layer.getRadius();
  }
  if (layer instanceof L.Rectangle) {
    return layer.getBounds().contains(L.latLng(lat, lng));
  }
  if (layer instanceof L.Polygon) {
    const latlngs = (layer.getLatLngs()[0] as any);
    let inside = false;
    for (let i = 0, j = latlngs.length - 1; i < latlngs.length; j = i++) {
      const xi = latlngs[i].lat, yi = latlngs[i].lng;
      const xj = latlngs[j].lat, yj = latlngs[j].lng;
      const intersect = ((yi > lng) !== (yj > lng)) &&
        (lat < (xj - xi) * (lng - yi) / (yj - yi) + xi);
      if (intersect) inside = !inside;
    }
    return inside;
  }
  return false;
}
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { db } from './firebase';
import { 
  collection, 
  addDoc, 
  getDocs, 
  deleteDoc, 
  doc, 
  query, 
  writeBatch,
  where,
  updateDoc
} from 'firebase/firestore';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const decodeHtml = (str: string): string => {
  if (!str) return str;
  return str
    .replace(/&#039;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&apos;/g, "'");
};

const normalizeDeposito = (name: string): string => {
  const decoded = decodeHtml(String(name || '').trim());
  const lower = decoded.toLowerCase();
  const targets = ['roma - pallavicini', 'roma - villa chigi', 'roma - sede', 'roma-sede'];
  if (targets.includes(lower)) {
    return 'Roma';
  }
  return decoded || 'N/D';
};

interface SaleRecord {
  audioprotesista: string;
  valore: number;
  data: Date;
}

interface BudgetDistribution {
  id?: string;
  year: number;
  percentages: number[]; // 12 numbers
}

interface PersonBudget {
  id?: string;
  year: number;
  name: string;
  annualBudget: number;
}

interface PersonnelOrder {
  id?: string;
  year: number;
  names: string[];
}

interface CreditNoteRecord {
  audioprotesista: string;
  valore: number;
  data: Date;
}

const MONTHS = [
  'Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
  'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'
];

function ComparisonTable({ data, years, onCellClick, comparisonType }: { data: any[], years: number[], onCellClick: any, comparisonType: 'audioprotesista' | 'recapiti' }) {
  if (data.length === 0) return (
    <div className="p-12 text-center opacity-40 italic font-serif text-lg border border-[#141414] border-dashed">
      Carica i dati per visualizzare le comparazioni
    </div>
  );

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-left border-collapse border border-[#141414]">
        <thead>
          <tr className="bg-gray-50 border-b border-[#141414]">
            <th className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 border-r border-[#141414] min-w-[130px]">Nome</th>
            {MONTHS.map((m, mIdx) => (
              <th key={mIdx} className="p-2 text-[10px] uppercase tracking-widest font-normal opacity-50 text-center border-r border-[#141414]" colSpan={3}>
                {m.slice(0, 3)}
              </th>
            ))}
            <th className="p-4 text-[10px] uppercase tracking-widest font-bold text-center bg-gray-100" colSpan={3}>Totale YTD</th>
          </tr>
          <tr className="bg-gray-100/50 border-b border-[#141414] text-[9px] uppercase tracking-tighter">
            <th className="border-r border-[#141414]"></th>
            {MONTHS.map((_, mIdx) => (
              <React.Fragment key={mIdx}>
                <th className="p-1 text-center border-r border-[#141414] text-gray-400">{years[2]}</th>
                <th className="p-1 text-center border-r border-[#141414] text-gray-600">{years[1]}</th>
                <th className="p-1 text-center border-r border-[#141414] bg-yellow-50 font-bold">{years[0]}</th>
              </React.Fragment>
            ))}
            <th className="p-1 text-center border-r border-[#141414] text-gray-400 bg-gray-100">{years[2]}</th>
            <th className="p-1 text-center border-r border-[#141414] text-gray-600 bg-gray-100">{years[1]}</th>
            <th className="p-1 text-center bg-yellow-100 font-bold">{years[0]}</th>
          </tr>
        </thead>
        <tbody className="font-mono text-[10px]">
          {(() => {
            if (comparisonType === 'recapiti') {
              const groups: Record<string, any[]> = {};
              data.forEach(row => {
                const dep = normalizeDeposito(row.deposito);
                if (!groups[dep]) groups[dep] = [];
                groups[dep].push(row);
              });
              
              const sortedGroups = Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0]));
              
              return sortedGroups.map(([deposito, items]) => {
                // Calculate group totals
                const groupTotals = years.map((_, yIdx) => {
                  const monthly = new Array(12).fill(0);
                  items.forEach(it => {
                    it.years[yIdx].monthly.forEach((val: number, mIdx: number) => {
                      monthly[mIdx] += val;
                    });
                  });
                  return {
                    monthly,
                    total: monthly.reduce((a, b) => a + b, 0)
                  };
                });

                return (
                  <React.Fragment key={deposito}>
                    <tr className="bg-gray-100 font-bold border-b border-[#141414]">
                      <td 
                        onClick={() => onCellClick(null, -1, undefined, undefined, comparisonType, deposito)}
                        className="p-2 pl-4 text-[9px] uppercase tracking-widest text-[#141414] cursor-pointer hover:bg-gray-200"
                        colSpan={1}
                      >
                        Deposito: {deposito}
                      </td>
                      {MONTHS.map((_, mIdx) => (
                        <React.Fragment key={mIdx}>
                          <td 
                            className="p-1 text-right border-r border-[#141414] bg-gray-100/50 cursor-pointer hover:bg-gray-200"
                            onClick={() => onCellClick(null, mIdx, undefined, years[2], comparisonType, deposito)}
                          >
                            {Math.round(groupTotals[2].monthly[mIdx]).toLocaleString('it-IT')}
                          </td>
                          <td 
                            className="p-1 text-right border-r border-[#141414] bg-gray-100/50 cursor-pointer hover:bg-gray-200"
                            onClick={() => onCellClick(null, mIdx, undefined, years[1], comparisonType, deposito)}
                          >
                            {Math.round(groupTotals[1].monthly[mIdx]).toLocaleString('it-IT')}
                          </td>
                          <td 
                            className="p-1 text-right border-r border-[#141414] bg-yellow-100/30 cursor-pointer hover:bg-yellow-200"
                            onClick={() => onCellClick(null, mIdx, undefined, years[0], comparisonType, deposito)}
                          >
                            {Math.round(groupTotals[0].monthly[mIdx]).toLocaleString('it-IT')}
                          </td>
                        </React.Fragment>
                      ))}
                      <td 
                        className="p-1 text-right border-r border-[#141414] bg-gray-200 cursor-pointer hover:bg-gray-300"
                        onClick={() => onCellClick(null, -1, undefined, years[2], comparisonType, deposito)}
                      >
                        {Math.round(groupTotals[2].total).toLocaleString('it-IT')}
                      </td>
                      <td 
                        className="p-1 text-right border-r border-[#141414] bg-gray-200 cursor-pointer hover:bg-gray-300"
                        onClick={() => onCellClick(null, -1, undefined, years[1], comparisonType, deposito)}
                      >
                        {Math.round(groupTotals[1].total).toLocaleString('it-IT')}
                      </td>
                      <td 
                        className="p-1 text-right bg-yellow-200 cursor-pointer hover:bg-yellow-300"
                        onClick={() => onCellClick(null, -1, undefined, years[0], comparisonType, deposito)}
                      >
                        {Math.round(groupTotals[0].total).toLocaleString('it-IT')}
                      </td>
                    </tr>
                    {items.map((row, idx) => (
                      <tr key={row.name} className={cn(
                        "border-b border-[#141414] hover:bg-gray-50 transition-colors",
                        idx % 2 === 0 ? "bg-white" : "bg-gray-50/20"
                      )}>
                        <td 
                          className="p-2 font-bold border-r border-[#141414] text-[11px] truncate pl-6 cursor-pointer hover:bg-gray-100 transition-colors" 
                          title={row.name}
                          onClick={() => onCellClick(row.name, -1, undefined, undefined, comparisonType)}
                        >
                          {row.name}
                        </td>
                        {MONTHS.map((_, mIdx) => {
                          const vals = row.years.map((y: any) => y.monthly[mIdx]);
                          return (
                            <React.Fragment key={mIdx}>
                              <td 
                                className="p-1 text-right border-r border-[#141414] opacity-70 cursor-pointer hover:bg-gray-100"
                                onClick={() => onCellClick(row.name, mIdx, undefined, years[2], comparisonType)}
                              >
                                {Math.round(vals[2]).toLocaleString('it-IT')}
                              </td>
                              <td 
                                className="p-1 text-right border-r border-[#141414] opacity-80 cursor-pointer hover:bg-gray-100"
                                onClick={() => onCellClick(row.name, mIdx, undefined, years[1], comparisonType)}
                              >
                                {Math.round(vals[1]).toLocaleString('it-IT')}
                              </td>
                              <td 
                                className="p-1 text-right border-r border-[#141414] font-bold cursor-pointer hover:bg-yellow-100 transition-colors bg-yellow-50/20"
                                onClick={() => onCellClick(row.name, mIdx, undefined, years[0], comparisonType)}
                              >
                                {Math.round(vals[0]).toLocaleString('it-IT')}
                              </td>
                            </React.Fragment>
                          );
                        })}
                        <td 
                          className="p-1 text-right border-r border-[#141414] font-bold opacity-70 bg-gray-50/50 cursor-pointer hover:bg-gray-200"
                          onClick={() => onCellClick(row.name, -1, undefined, years[2], comparisonType)}
                        >
                          {Math.round(row.years[2].total).toLocaleString('it-IT')}
                        </td>
                        <td 
                          className="p-1 text-right border-r border-[#141414] font-bold opacity-80 bg-gray-50/50 cursor-pointer hover:bg-gray-200"
                          onClick={() => onCellClick(row.name, -1, undefined, years[1], comparisonType)}
                        >
                          {Math.round(row.years[1].total).toLocaleString('it-IT')}
                        </td>
                        <td 
                          className="p-1 text-right font-bold bg-yellow-50 cursor-pointer hover:bg-yellow-200"
                          onClick={() => onCellClick(row.name, -1, undefined, years[0], comparisonType)}
                        >
                          {Math.round(row.years[0].total).toLocaleString('it-IT')}
                        </td>
                      </tr>
                    ))}
                  </React.Fragment>
                );
              });
            }
            
            return data.map((row, idx) => (
              <tr key={row.name} className={cn(
                "border-b border-[#141414] hover:bg-gray-50 transition-colors",
                idx % 2 === 0 ? "bg-white" : "bg-gray-50/20"
              )}>
                <td 
                  className="p-2 font-bold border-r border-[#141414] text-[11px] truncate cursor-pointer hover:bg-gray-100 transition-colors" 
                  title={row.name}
                  onClick={() => onCellClick(row.name, -1, undefined, undefined, comparisonType)}
                >
                  {row.name}
                </td>
                {MONTHS.map((_, mIdx) => {
                  const vals = row.years.map((y: any) => y.monthly[mIdx]);
                  return (
                    <React.Fragment key={mIdx}>
                      <td 
                        className="p-1 text-right border-r border-[#141414] opacity-70 cursor-pointer hover:bg-gray-100"
                        onClick={() => onCellClick(row.name, mIdx, undefined, years[2], comparisonType)}
                      >
                        {Math.round(vals[2]).toLocaleString('it-IT')}
                      </td>
                      <td 
                        className="p-1 text-right border-r border-[#141414] opacity-80 cursor-pointer hover:bg-gray-100"
                        onClick={() => onCellClick(row.name, mIdx, undefined, years[1], comparisonType)}
                      >
                        {Math.round(vals[1]).toLocaleString('it-IT')}
                      </td>
                      <td 
                        className="p-1 text-right border-r border-[#141414] font-bold cursor-pointer hover:bg-yellow-100 transition-colors bg-yellow-50/20"
                        onClick={() => onCellClick(row.name, mIdx, undefined, years[0], comparisonType)}
                      >
                        {Math.round(vals[0]).toLocaleString('it-IT')}
                      </td>
                    </React.Fragment>
                  );
                })}
                <td 
                  className="p-1 text-right border-r border-[#141414] font-bold opacity-70 bg-gray-50/50 cursor-pointer hover:bg-gray-200"
                  onClick={() => onCellClick(row.name, -1, undefined, years[2], comparisonType)}
                >
                  {Math.round(row.years[2].total).toLocaleString('it-IT')}
                </td>
                <td 
                  className="p-1 text-right border-r border-[#141414] font-bold opacity-80 bg-gray-50/50 cursor-pointer hover:bg-gray-200"
                  onClick={() => onCellClick(row.name, -1, undefined, years[1], comparisonType)}
                >
                  {Math.round(row.years[1].total).toLocaleString('it-IT')}
                </td>
                <td 
                  className="p-1 text-right font-bold bg-yellow-50 cursor-pointer hover:bg-yellow-200"
                  onClick={() => onCellClick(row.name, -1, undefined, years[0], comparisonType)}
                >
                  {Math.round(row.years[0].total).toLocaleString('it-IT')}
                </td>
              </tr>
            ));
          })()}
        </tbody>
      </table>
    </div>
  );
}

export default function App() {
  const [activeTab, setActiveTab] = useState<'sales' | 'patients' | 'appointmentsManagement' | 'screeningEntrances' | 'noah' | 'comparisons' | 'cap' | 'recapiti' | 'verifiche' | 'channels' | 'audioAnalysis' | 'openTrials' | 'budget'>('patients');

  const [salesData, setSalesData] = useState<any[]>([]);
  const [creditNotes, setCreditNotes] = useState<any[]>([]);
  const [trialsData, setTrialsData] = useState<any[]>([]);

  // Fetch unique audioprotesisti from data
  const allAudioprotesisti = useMemo(() => {
    const names = new Set<string>();
    salesData.forEach(s => s.audioprotesista && names.add(s.audioprotesista));
    creditNotes.forEach(c => c.audioprotesista && names.add(c.audioprotesista));
    trialsData.forEach(t => t.audioprotesista && names.add(t.audioprotesista));
    return Array.from(names).sort();
  }, [salesData, creditNotes, trialsData]);

  // Process data for Open Trials section
  const openTrialsAnalysis = useMemo(() => {
    if (trialsData.length === 0) return [];

    // Filter for "Prova Aperta"
    const openTrials = trialsData.filter(t => 
      String(t.statoDocumento || '').toLowerCase().trim() === 'prova aperta'
    );

    // Group by Audioprotesista
    const grouped: Record<string, any[]> = {};
    const today = new Date();

    openTrials.forEach(t => {
      const name = t.audioprotesista || 'N/D';
      if (!grouped[name]) grouped[name] = [];

      // Calculate days elapsed
      let daysElapsed = 0;
      if (t.fullDate) {
        const parts = t.fullDate.split('/');
        if (parts.length === 3) {
          const docDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
          const diffTime = Math.abs(today.getTime() - docDate.getTime());
          daysElapsed = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        }
      }

      grouped[name].push({
        ...t,
        daysElapsed
      });
    });

    return Object.entries(grouped).sort((a, b) => a[0].localeCompare(b[0]));
  }, [trialsData]);
  const [selectedChannels, setSelectedChannels] = useState<string[]>([]);
  const [showChannelDropdown, setShowChannelDropdown] = useState(false);
  
  // Comparison years state
  const [comparisonAnchorYear, setComparisonAnchorYear] = useState<number>(new Date().getFullYear());
  
  // New states for patients management
  const [patientsData, setPatientsData] = useState<any[]>([]);
  const [callsData, setCallsData] = useState<any[]>([]);
  const [appointmentsData, setAppointmentsData] = useState<any[]>([]);
  const [storesData, setStoresData] = useState<any[]>([]);
  const [storeCaps, setStoreCaps] = useState<any[]>([]);
  const [analyzedPatients, setAnalyzedPatients] = useState<any[]>([]);
  const [noahPatients, setNoahPatients] = useState<any[]>([]);
  const [noahThreshold, setNoahThreshold] = useState<number>(35);
  const [marketingAnalysis, setMarketingAnalysis] = useState<any[]>([]);
  const [screeningAnalysis, setScreeningAnalysis] = useState<any[]>([]);
  const [showAllDetails, setShowAllDetails] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [showMarketingAnalysis, setShowMarketingAnalysis] = useState(false);
  const [showScreeningAnalysis, setShowScreeningAnalysis] = useState(false);
  const [showCoefficientiAnalysis, setShowCoefficientiAnalysis] = useState(false);
  const [coefficientiAnalysis, setCoefficientiAnalysis] = useState<any[]>([]);
  const [showMediciAnalysis, setShowMediciAnalysis] = useState(false);
  const [mediciAnalysis, setMediciAnalysis] = useState<any[]>([]);
  const [selectedBranch, setSelectedBranch] = useState<string>('');
  const [selectedStoreId, setSelectedStoreId] = useState<string>('');
  const [selectedPatientType, setSelectedPatientType] = useState<string>('');
  const [selectedCap, setSelectedCap] = useState<string>('');
  const [storeSearch, setStoreSearch] = useState<string>('');
  const [capFilterSearch, setCapFilterSearch] = useState<string>('');
  const [showStoreDropdown, setShowStoreDropdown] = useState<boolean>(false);
  const [showCapFilterDropdown, setShowCapFilterDropdown] = useState<boolean>(false);
  const [selectedMonths, setSelectedMonths] = useState<number>(3);
  const [selectedAudioPro, setSelectedAudioPro] = useState<string>('');
  const [selectedAudioYear, setSelectedAudioYear] = useState<number>(new Date().getFullYear());
  const [visibleDepositi, setVisibleDepositi] = useState<string[]>([]);
  const [capPage, setCapPage] = useState(1);
  const [capInput, setCapInput] = useState('');
  const [capSearch, setCapSearch] = useState('');
  const [capSortConfig, setCapSortConfig] = useState<{ key: string, direction: 'asc' | 'desc' } | null>({ key: 'citta', direction: 'asc' });
  const [selectedStoreIds, setSelectedStoreIds] = useState<string[]>([]);
  const [selectedRecapitiIds, setSelectedRecapitiIds] = useState<string[]>([]);
  const [recapitiDates, setRecapitiDates] = useState<Record<string, string>>({});
  const [availabilityData, setAvailabilityData] = useState<any[]>([]);
  const [availabilityMonth, setAvailabilityMonth] = useState<'current' | 'next'>('current');
  const [verificationResults, setVerificationResults] = useState<any[]>([]);
  const [activeVerification, setActiveVerification] = useState<string | null>(null);
  const [showMap, setShowMap] = useState(false);
  const [geocodedPatients, setGeocodedPatients] = useState<any[]>([]);
  const [geocodedStores, setGeocodedStores] = useState<any[]>([]);
  const [isGeocoding, setIsGeocoding] = useState(false);
  const [geocodingProgress, setGeocodingProgress] = useState({ current: 0, total: 0 });
  const itemsPerPage = 15;

  const MONTH_NAMES = ["Gen", "Feb", "Mar", "Apr", "Mag", "Giu", "Lug", "Ago", "Set", "Ott", "Nov", "Dic"];

  const [selectedYear, setSelectedYear] = useState<number>(new Date().getFullYear());
  const [isBudgetSummaryOpen, setIsBudgetSummaryOpen] = useState(false);
  const [appointmentsAnalysisView, setAppointmentsAnalysisView] = useState<'canali' | 'tipi'>('canali');

  // Budget states
  const [monthlyDistribution, setMonthlyDistribution] = useState<number[]>(new Array(12).fill(0).map((_, i) => [8, 7, 9, 8, 9, 8, 7, 6, 9, 10, 10, 9][i])); // Default values
  const [personnelBudgets, setPersonnelBudgets] = useState<PersonBudget[]>([]);
  const [personnelOrder, setPersonnelOrder] = useState<string[]>([]);
  const [isSavingBudget, setIsSavingBudget] = useState(false);

  // Load Budget and Order from Firestore
  useEffect(() => {
    const loadBudgets = async () => {
      try {
        const distQuery = query(collection(db, 'budgetDistributions'), where('year', '==', selectedYear));
        const distDocs = await getDocs(distQuery);
        if (!distDocs.empty) {
          const data = distDocs.docs[0].data() as BudgetDistribution;
          if (data.percentages) setMonthlyDistribution(data.percentages);
        } else {
          setMonthlyDistribution([8, 7, 9, 8, 9, 8, 7, 6, 9, 10, 10, 9]);
        }

        const personQuery = query(collection(db, 'personnelBudgets'), where('year', '==', selectedYear));
        const personDocs = await getDocs(personQuery);
        const budgets: PersonBudget[] = personDocs.docs.map(d => ({ id: d.id, ...d.data() } as PersonBudget));
        setPersonnelBudgets(budgets);

        const orderQuery = query(collection(db, 'personnelOrders'), where('year', '==', selectedYear));
        const orderDocs = await getDocs(orderQuery);
        if (!orderDocs.empty) {
          const orderData = orderDocs.docs[0].data() as PersonnelOrder;
          setPersonnelOrder(orderData.names || []);
        } else {
          setPersonnelOrder([]);
        }
      } catch (err) {
        console.error("Error loading budgets:", err);
      }
    };
    loadBudgets();
  }, [selectedYear]);

  // Combined personnel list: ordered items first, then newcomers
  const orderedPersonnelList = useMemo(() => {
    const combined = [...personnelOrder];
    allAudioprotesisti.forEach(name => {
      if (!combined.includes(name)) {
        combined.push(name);
      }
    });
    // Filter out names that no longer exist in data
    return combined.filter(name => allAudioprotesisti.includes(name));
  }, [allAudioprotesisti, personnelOrder]);

  const savePersonnelOrder = async (newOrder: string[]) => {
    try {
      const q = query(collection(db, 'personnelOrders'), where('year', '==', selectedYear));
      const snapshot = await getDocs(q);
      if (snapshot.empty) {
        await addDoc(collection(db, 'personnelOrders'), { year: selectedYear, names: newOrder });
      } else {
        const docRef = doc(db, 'personnelOrders', snapshot.docs[0].id);
        await updateDoc(docRef, { names: newOrder });
      }
      setPersonnelOrder(newOrder);
    } catch (err) {
      console.error("Error saving order:", err);
    }
  };

  const movePersonnel = async (index: number, direction: 'up' | 'down') => {
    const newOrder = [...orderedPersonnelList];
    const targetIndex = direction === 'up' ? index - 1 : index + 1;
    if (targetIndex < 0 || targetIndex >= newOrder.length) return;

    [newOrder[index], newOrder[targetIndex]] = [newOrder[targetIndex], newOrder[index]];
    setPersonnelOrder(newOrder);
    await savePersonnelOrder(newOrder);
  };

  const saveMonthlyDistribution = async (newPercentages: number[]) => {
    setIsSavingBudget(true);
    try {
      const q = query(collection(db, 'budgetDistributions'), where('year', '==', selectedYear));
      const snapshot = await getDocs(q);
      if (snapshot.empty) {
        await addDoc(collection(db, 'budgetDistributions'), { year: selectedYear, percentages: newPercentages });
      } else {
        const docRef = doc(db, 'budgetDistributions', snapshot.docs[0].id);
        await updateDoc(docRef, { percentages: newPercentages });
      }
      setMonthlyDistribution(newPercentages);
    } catch (err) {
      console.error("Error saving distribution:", err);
    } finally {
      setIsSavingBudget(false);
    }
  };

  const savePersonBudget = async (name: string, annualBudget: number) => {
    setIsSavingBudget(true);
    try {
      const q = query(collection(db, 'personnelBudgets'), where('year', '==', selectedYear), where('name', '==', name));
      const snapshot = await getDocs(q);
      
      if (!snapshot.empty) {
        const docRef = doc(db, 'personnelBudgets', snapshot.docs[0].id);
        await updateDoc(docRef, { annualBudget });
        setPersonnelBudgets(prev => prev.map(p => p.name === name ? { ...p, annualBudget } : p));
      } else {
        const newDoc = await addDoc(collection(db, 'personnelBudgets'), { year: selectedYear, name, annualBudget });
        setPersonnelBudgets(prev => [...prev, { id: newDoc.id, year: selectedYear, name, annualBudget }]);
      }
    } catch (err) {
      console.error("Error saving person budget:", err);
    } finally {
      setIsSavingBudget(false);
    }
  };
  const [selectedChannelYear, setSelectedChannelYear] = useState<number>(new Date().getFullYear());
  const [drillDown, setDrillDown] = useState<{
    name: string;
    month: number;
    channel?: string;
    year?: number;
    sales: any[];
    credits: any[];
  } | null>(null);
  const [channelDrillDown, setChannelDrillDown] = useState<{
    title: string;
    subtitle: string;
    data: any[];
    type: 'patients' | 'calls' | 'appointments' | 'trials' | 'sales' | 'availability';
  } | null>(null);
  const [loading, setLoading] = useState({ 
    sales: false, 
    credits: false, 
    trials: false,
    patients: false, 
    calls: false, 
    appointments: false, 
    noah: false,
    stores: false,
    availability: false
  });
  const [error, setError] = useState<string | null>(null);

  const [selectedApptYear, setSelectedApptYear] = useState<number>(new Date().getFullYear());
  const [selectedScreeningYear, setSelectedScreeningYear] = useState<number>(new Date().getFullYear());
  
  const availableYears = useMemo(() => {
    const currentYear = new Date().getFullYear();
    const years = new Set<number>();
    
    // Always include 2022 to next year
    for (let y = 2022; y <= currentYear + 1; y++) {
      years.add(y);
    }

    salesData.forEach(s => years.add(s.year));
    creditNotes.forEach(c => years.add(c.year));
    trialsData.forEach(t => years.add(t.year));
    const sorted = Array.from(years).sort((a, b) => b - a);
    return sorted;
  }, [salesData, creditNotes, trialsData]);

  const availableChannelYears = useMemo(() => {
    const years = new Set<number>();
    patientsData.forEach(p => { if (p.year) years.add(p.year); });
    const sorted = Array.from(years).sort((a, b) => b - a);
    return sorted.length > 0 ? sorted : [new Date().getFullYear()];
  }, [patientsData]);

  const availableApptYears = useMemo(() => {
    const years = new Set<number>();
    appointmentsData.forEach(a => {
      if (a.year) years.add(a.year);
    });
    const sorted = Array.from(years).sort((a, b) => b - a);
    return sorted.length > 0 ? sorted : [new Date().getFullYear()];
  }, [appointmentsData]);

  const appointmentsAnalysis = useMemo(() => {
    if (appointmentsData.length === 0) return [];

    // Filter by year and requested conditions
    // models/appointments.fields.appointment_result = "Si è presentato"
    // Tipo = "controllo udito" or "Prima Visita"
    const filtered = appointmentsData.filter(a => {
      if (a.year !== selectedApptYear) return false;
      
      const esito = String(a.esito || '').toLowerCase().trim();
      const tipo = String(a.tipo || '').toLowerCase().trim();
      
      const matchEsito = esito === 'si è presentato' || esito === 'si e presentato' || esito === 'presentato';
      const matchTipo = tipo === 'controllo udito' || tipo === 'prima visita';
      
      return matchEsito && matchTipo;
    });

    // Group by Audioprotesista, then by selected subGroup
    const stats = new Map<string, {
      name: string,
      months: number[], // Count per month
      subGroups: Map<string, number[]> // Channel or Type -> Month counts
    }>();

    filtered.forEach(a => {
      const pro = a.audioprotesista || 'Da assegnare';
      const subGroupValue = appointmentsAnalysisView === 'canali' ? (a.canale || 'N/D') : (a.contactType || 'N/D');
      const month = a.month;

      if (!stats.has(pro)) {
        stats.set(pro, {
          name: pro,
          months: new Array(12).fill(0),
          subGroups: new Map()
        });
      }

      const proStats = stats.get(pro)!;
      proStats.months[month]++;

      if (!proStats.subGroups.has(subGroupValue)) {
        proStats.subGroups.set(subGroupValue, new Array(12).fill(0));
      }
      proStats.subGroups.get(subGroupValue)![month]++;
    });

    return Array.from(stats.values()).sort((a, b) => a.name.localeCompare(b.name));
  }, [appointmentsData, selectedApptYear, appointmentsAnalysisView]);

  const screeningEntrancesAnalysis = useMemo(() => {
    if (appointmentsData.length === 0) return [];

    const excludedStores = ['genova', 'domicilio - genova', 'fabriano', 'roma - pallavicini', 'domicilio - roma pallavicini', 'civitanova marche', 'viareggio', 'aprilia'];

    const filtered = appointmentsData.filter(a => {
      if (a.year !== selectedScreeningYear) return false;
      
      const esito = String(a.esito || '').toLowerCase().trim();
      const tipo = String(a.tipo || '').toLowerCase().trim();
      
      const matchEsito = esito === 'si è presentato' || esito === 'si e presentato' || esito === 'presentato';
      const matchTipo = tipo === 'controllo udito' || tipo === 'prima visita';
      
      const store = String(a.sede || '').toLowerCase().trim();
      if (excludedStores.includes(store)) return false;

      return matchEsito && matchTipo;
    });

    const stats = new Map<string, {
      name: string,
      months: number[],
      stores: Map<string, number[]> // Store -> Month counts
    }>();

    filtered.forEach(a => {
      const pro = a.audioprotesista || 'Da assegnare';
      const store = a.sede || 'N/D';
      const month = a.month;

      if (!stats.has(pro)) {
        stats.set(pro, {
          name: pro,
          months: new Array(12).fill(0),
          stores: new Map()
        });
      }

      const proStats = stats.get(pro)!;
      proStats.months[month]++;

      if (!proStats.stores.has(store)) {
        proStats.stores.set(store, new Array(12).fill(0));
      }
      proStats.stores.get(store)![month]++;
    });

    return Array.from(stats.values()).sort((a, b) => a.name.localeCompare(b.name));
  }, [appointmentsData, selectedScreeningYear]);

  const budgetSituationData = useMemo(() => {
    // Group monthly actuals by person
    const monthlyActuals = new Map<string, number[]>();
    
    salesData.forEach(s => {
      if (s.year === selectedYear) {
        if (!monthlyActuals.has(s.audioprotesista)) {
          monthlyActuals.set(s.audioprotesista, new Array(12).fill(0));
        }
        monthlyActuals.get(s.audioprotesista)![s.month] += (s.valore || 0);
      }
    });

    creditNotes.forEach(c => {
      if (c.year === selectedYear) {
        if (!monthlyActuals.has(c.audioprotesista)) {
          monthlyActuals.set(c.audioprotesista, new Array(12).fill(0));
        }
        monthlyActuals.get(c.audioprotesista)![c.month] -= (c.valore || 0);
      }
    });

    return orderedPersonnelList.map(name => {
      const personBudget = personnelBudgets.find(b => b.name === name);
      const annualBudget = personBudget?.annualBudget || 0;
      const actuals = monthlyActuals.get(name) || new Array(12).fill(0);
      
      const months = actuals.map((actual, mIdx) => {
        const budget = (annualBudget * (monthlyDistribution[mIdx] || 0)) / 100;
        return {
          month: mIdx,
          actual,
          budget,
          diff: actual - budget
        };
      });

      return {
        name,
        months,
        totalAnnualBudget: annualBudget,
        totalAnnualActual: actuals.reduce((a, b) => a + b, 0)
      };
    });
  }, [salesData, creditNotes, selectedYear, orderedPersonnelList, personnelBudgets, monthlyDistribution]);

  const audioProList = useMemo(() => {
    const names = new Set<string>();
    [availabilityData, callsData, appointmentsData, salesData, creditNotes, trialsData].forEach(dataset => {
      dataset.forEach(d => {
        if (d.audioprotesista && d.audioprotesista !== 'N/D') {
          names.add(d.audioprotesista);
        }
      });
    });
    return Array.from(names).sort();
  }, [availabilityData, callsData, appointmentsData, salesData, creditNotes, trialsData]);

  const audioPerformance = useMemo(() => {
    if (!selectedAudioPro || !selectedAudioYear) return null;

    // 1. Events (dd.xlsx)
    const events = availabilityData.filter(d => 
      String(d.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
      d.year === selectedAudioYear &&
      !['domicilio - fabriano', 'fabriano', 'domicilio - genova', 'domicilio', 'genova', 'domicilio - roma pallavicini', 'roma - pallavicini', 'aprilia', 'ricerca struttura'].includes(String(d.storeName || '').toLowerCase().trim())
    );
    const eventsByStore: Record<string, number[]> = {};
    events.forEach(e => {
      if (!eventsByStore[e.storeName]) eventsByStore[e.storeName] = new Array(12).fill(0);
      eventsByStore[e.storeName][e.month]++;
    });

    // 2. Phone Calls (tt.xlsx)
    const calls = callsData.filter(c => 
      String(c.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
      c.year === selectedAudioYear
    );
    const callsByOutcome: Record<string, number[]> = {};
    calls.forEach(c => {
      const outcome = c.contatto || 'N/D';
      if (!callsByOutcome[outcome]) callsByOutcome[outcome] = new Array(12).fill(0);
      callsByOutcome[outcome][c.month]++;
    });

    // 3. Appointments (aa.xlsx)
    const appts = appointmentsData.filter(a => {
      const aType = String(a.tipo || '').toLowerCase().trim();
      const aContactType = String(a.contactType || '').toLowerCase().trim();
      const isCorrectType = aType.includes('controllo udito') || aType.includes('prima visita') || 
                            aContactType.includes('controllo udito') || aContactType.includes('prima visita');
      
      const esito = String(a.esito || '').toLowerCase().trim();
      const isPresentato = esito === 'si è presentato' || esito === 'si e presentato';

      return String(a.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
        a.year === selectedAudioYear &&
        isPresentato &&
        isCorrectType;
    });
    const apptsByType: Record<string, number[]> = {};
    appts.forEach(a => {
      // Use contactType for grouping (rows) as requested
      const type = a.contactType !== 'N/D' ? a.contactType : (a.tipo !== 'N/D' ? a.tipo : 'N/D');
      if (!apptsByType[type]) apptsByType[type] = new Array(12).fill(0);
      apptsByType[type][a.month]++;
    });

    return { 
      events: Object.entries(eventsByStore).sort((a, b) => a[0].localeCompare(b[0])),
      calls: Object.entries(callsByOutcome).sort((a, b) => a[0].localeCompare(b[0])),
      appointments: Object.entries(apptsByType).sort((a, b) => a[0].localeCompare(b[0])),
      openTrials: trialsData.filter(t => 
        String(t.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
        t.year === selectedAudioYear &&
        String(t.statoDocumento || '').toLowerCase().trim() === 'prova aperta'
      ).sort((a, b) => (a.cliente || '').localeCompare(b.cliente || '')),
      finance: (() => {
        const audioSales = salesData.filter(s => 
          String(s.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
          s.year === selectedAudioYear
        );
        const audioCredits = creditNotes.filter(c => 
          String(c.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
          c.year === selectedAudioYear
        );

        const monthlySales = new Array(12).fill(0);
        const monthlyCredits = new Array(12).fill(0);

        audioSales.forEach(s => monthlySales[s.month] += (s.valore || 0));
        audioCredits.forEach(c => monthlyCredits[c.month] += (c.valore || 0));

        return {
          sales: monthlySales,
          credits: monthlyCredits,
          balance: monthlySales.map((v, i) => v - monthlyCredits[i])
        };
      })()
    };
  }, [selectedAudioPro, selectedAudioYear, availabilityData, callsData, appointmentsData, trialsData, salesData, creditNotes]);

  const availableChannels = useMemo(() => {
    const channels = new Set<string>();
    patientsData.forEach(p => p.canale && p.canale !== 'N/D' && channels.add(p.canale));
    callsData.forEach(c => c.canale && c.canale !== 'N/D' && channels.add(c.canale));
    appointmentsData.forEach(a => a.canale && a.canale !== 'N/D' && channels.add(a.canale));
    trialsData.forEach(t => t.canale && t.canale !== 'N/D' && channels.add(t.canale));
    salesData.forEach(s => s.canale && s.canale !== 'N/D' && channels.add(s.canale));
    return Array.from(channels).sort();
  }, [patientsData, callsData, appointmentsData, trialsData, salesData]);

  const channelAnalysis = useMemo(() => {
    if (selectedChannels.length === 0) return null;

    const filterByChannelAndYear = (data: any[]) => 
      data.filter(item => 
        selectedChannels.some(sc => String(item.canale || '').toLowerCase().trim() === sc.toLowerCase().trim()) &&
        item.year === selectedChannelYear
      );

    const filteredPatients = filterByChannelAndYear(patientsData);
    const filteredCalls = filterByChannelAndYear(callsData);
    const filteredAppointments = filterByChannelAndYear(appointmentsData);
    const filteredTrials = filterByChannelAndYear(trialsData);
    const filteredSales = filterByChannelAndYear(salesData);

    // 1. Anagrafiche (by month and store)
    const anagraficheByStore: Record<string, Record<number, number>> = {};
    const months = new Set<number>();
    filteredPatients.forEach(p => {
      const store = p.store || 'N/D';
      const month = (p.month || 0) + 1; // 1-indexed
      months.add(month);
      if (!anagraficheByStore[store]) anagraficheByStore[store] = {};
      anagraficheByStore[store][month] = (anagraficheByStore[store][month] || 0) + 1;
    });

    // 2. Telefonate (by outcome)
    const callsByOutcome: Record<string, number> = {};
    filteredCalls.forEach(c => {
      const outcome = c.contatto || 'N/D';
      callsByOutcome[outcome] = (callsByOutcome[outcome] || 0) + 1;
    });

    // 3. Appuntamenti (by status and month) - Filtered by "Prima Visita"
    const appointmentsByStatus: Record<string, Record<number, number>> = {};
    const apptMonths = new Set<number>();
    filteredAppointments
      .filter(a => String(a.tipo || '').toLowerCase().trim() === 'prima visita')
      .forEach(a => {
        const status = a.contatto || 'N/D';
        const month = (a.month || 0) + 1;
        apptMonths.add(month);
        if (!appointmentsByStatus[status]) appointmentsByStatus[status] = {};
        appointmentsByStatus[status][month] = (appointmentsByStatus[status][month] || 0) + 1;
      });

    // 4. Presentati (by contact type and month) - Filtered by "Prima Visita" and "Si è presentato"
    const presentatiByType: Record<string, Record<number, number>> = {};
    const presMonths = new Set<number>();
    filteredAppointments
      .filter(a => {
        const tipo = String(a.tipo || '').toLowerCase().trim();
        const status = String(a.contatto || '').toLowerCase().trim();
        return tipo === 'prima visita' && status === 'si è presentato';
      })
      .forEach(a => {
        const type = a.contactType || 'N/D';
        const month = (a.month || 0) + 1;
        presMonths.add(month);
        if (!presentatiByType[type]) presentatiByType[type] = {};
        presentatiByType[type][month] = (presentatiByType[type][month] || 0) + 1;
      });

    // 5. Prove (by status and month) - Using statoDocumento
    const trialsByStatus: Record<string, Record<number, number>> = {};
    const trialMonths = new Set<number>();
    filteredTrials.forEach(t => {
      const status = t.statoDocumento || 'N/D';
      const month = (t.month || 0) + 1;
      trialMonths.add(month);
      if (!trialsByStatus[status]) trialsByStatus[status] = {};
      trialsByStatus[status][month] = (trialsByStatus[status][month] || 0) + 1;
    });

    // 6. Fatturato (by salesperson and month)
    const salesByPerson: Record<string, Record<number, number>> = {};
    const salesMonths = new Set<number>();
    filteredSales.forEach(s => {
      const person = s.audioprotesista || 'N/D';
      const month = (s.month || 0) + 1;
      salesMonths.add(month);
      if (!salesByPerson[person]) salesByPerson[person] = {};
      salesByPerson[person][month] = (salesByPerson[person][month] || 0) + (s.valore || 0);
    });

    return {
      anagrafiche: {
        data: anagraficheByStore,
        months: Array.from(months).sort((a, b) => a - b)
      },
      calls: callsByOutcome,
      appointments: {
        data: appointmentsByStatus,
        months: Array.from(apptMonths).sort((a, b) => a - b)
      },
      presentati: {
        data: presentatiByType,
        months: Array.from(presMonths).sort((a, b) => a - b)
      },
      trials: {
        data: trialsByStatus,
        months: Array.from(trialMonths).sort((a, b) => a - b)
      },
      sales: {
        data: salesByPerson,
        months: Array.from(salesMonths).sort((a, b) => a - b)
      },
      filteredPatients,
      filteredCalls,
      filteredAppointments,
      filteredTrials,
      filteredSales
    };
  }, [selectedChannels, selectedChannelYear, patientsData, callsData, appointmentsData, trialsData, salesData]);

  const handleDownloadAnalysis = async () => {
    if (!channelAnalysis) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Analisi Canali');

    // Helper for styling
    const applyTableStyles = (startRow: number, startCol: number, title: string, headers: string[], data: any[][], hasFooter = false) => {
      // Title
      const titleCell = sheet.getCell(startRow, startCol);
      titleCell.value = title;
      titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
      titleCell.font = { bold: true, size: 12 };
      titleCell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

      // Headers
      headers.forEach((h, i) => {
        const cell = sheet.getCell(startRow + 1, startCol + i);
        cell.value = h;
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E0E0E0' } };
        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
      });

      // Data
      data.forEach((row, rowIndex) => {
        row.forEach((val, colIndex) => {
          const cell = sheet.getCell(startRow + 2 + rowIndex, startCol + colIndex);
          cell.value = val;
          cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
          
          // Currency formatting for sales
          if (title.toLowerCase().includes('fatturato') && colIndex > 0) {
             cell.numFmt = '"€" #,##0.00';
          }
        });
        
        // Highlight footer
        if (hasFooter && rowIndex === data.length - 1) {
          row.forEach((_, colIndex) => {
            const cell = sheet.getCell(startRow + 2 + rowIndex, startCol + colIndex);
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F2F2' } };
          });
        }
      });

      return startRow + data.length + 4; // Return next row
    };

    // Metadata
    sheet.getCell('A1').value = 'Report Analisi Canali';
    sheet.getCell('A1').font = { bold: true, size: 16 };
    sheet.getCell('A2').value = `Anno: ${selectedYear}`;
    sheet.getCell('A3').value = `Canali: ${selectedChannels.join(', ')}`;

    // 1. Anagrafiche
    const anagraficheHeaders = ['Store', ...channelAnalysis.anagrafiche.months.map(m => MONTH_NAMES[m - 1]), 'Totale'];
    const anagraficheRows: any[][] = [];
    Object.entries(channelAnalysis.anagrafiche.data).forEach(([store, months]) => {
      const row: (string | number)[] = [store];
      let total = 0;
      channelAnalysis.anagrafiche.months.forEach(m => {
        const val = (months as any)[m] || 0;
        row.push(val);
        total += val;
      });
      row.push(total);
      anagraficheRows.push(row);
    });
    const anagraficheTotals: (string | number)[] = ['Totale complessivo'];
    let grandTotalAnagrafiche = 0;
    channelAnalysis.anagrafiche.months.forEach(m => {
      const monthTotal = Object.values(channelAnalysis.anagrafiche.data).reduce((sum, months) => 
        (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
      anagraficheTotals.push(monthTotal as number);
      grandTotalAnagrafiche += monthTotal as number;
    });
    anagraficheTotals.push(grandTotalAnagrafiche);
    anagraficheRows.push(anagraficheTotals);

    const nextRowAnagrafiche = applyTableStyles(5, 1, 'Anagrafiche', anagraficheHeaders, anagraficheRows, true);

    // 2. Telefonate
    const callsHeaders = ['Esito', 'Conteggio'];
    const callsRows = Object.entries(channelAnalysis.calls).map(([outcome, count]) => [outcome, count]);
    const callsTotal = Object.values(channelAnalysis.calls).reduce((a, b) => (a as number) + (b as number), 0);
    callsRows.push(['Totale complessivo', callsTotal]);
    const nextRowTelefonate = applyTableStyles(nextRowAnagrafiche, 1, 'Telefonate', callsHeaders, callsRows, true);

    // 3. Appuntamenti (under Telefonate)
    const apptsHeaders = ['Stato', ...(channelAnalysis.appointments as any).months.map((m: number) => MONTH_NAMES[m - 1]), 'Totale'];
    const apptsRows: any[][] = [];
    Object.entries((channelAnalysis.appointments as any).data).forEach(([status, months]) => {
      const row: (string | number)[] = [status];
      let total = 0;
      (channelAnalysis.appointments as any).months.forEach((m: number) => {
        const val = (months as any)[m] || 0;
        row.push(val);
        total += val;
      });
      row.push(total);
      apptsRows.push(row);
    });
    const apptsTotals: (string | number)[] = ['Totale complessivo'];
    let grandTotalAppts = 0;
    (channelAnalysis.appointments as any).months.forEach((m: number) => {
      const monthTotal = Object.values((channelAnalysis.appointments as any).data).reduce((sum, months) => 
        (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
      apptsTotals.push(monthTotal as number);
      grandTotalAppts += monthTotal as number;
    });
    apptsTotals.push(grandTotalAppts);
    apptsRows.push(apptsTotals);
    const nextRowAppuntamenti = applyTableStyles(nextRowTelefonate, 1, 'Appuntamenti', apptsHeaders, apptsRows, true);

    // 4. Presentati (under Appuntamenti)
    const presentatiHeaders = ['Tipo Contatto', ...(channelAnalysis.presentati as any).months.map((m: number) => MONTH_NAMES[m - 1]), 'Totale'];
    const presentatiRows: any[][] = [];
    Object.entries((channelAnalysis.presentati as any).data).forEach(([type, months]) => {
      const row: (string | number)[] = [type];
      let total = 0;
      (channelAnalysis.presentati as any).months.forEach((m: number) => {
        const val = (months as any)[m] || 0;
        row.push(val);
        total += val;
      });
      row.push(total);
      presentatiRows.push(row);
    });
    const presentatiTotals: (string | number)[] = ['Totale complessivo'];
    let grandTotalPresentati = 0;
    (channelAnalysis.presentati as any).months.forEach((m: number) => {
      const monthTotal = Object.values((channelAnalysis.presentati as any).data).reduce((sum, months) => 
        (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
      presentatiTotals.push(monthTotal as number);
      grandTotalPresentati += monthTotal as number;
    });
    presentatiTotals.push(grandTotalPresentati);
    presentatiRows.push(presentatiTotals);
    const nextRowPresentati = applyTableStyles(nextRowAppuntamenti, 1, 'Presentati', presentatiHeaders, presentatiRows, true);

    // 5. Prove (under Presentati)
    const trialsHeaders = ['Stato documento', ...(channelAnalysis.trials as any).months.map((m: number) => MONTH_NAMES[m - 1]), 'Totale'];
    const trialsRows: any[][] = [];
    Object.entries((channelAnalysis.trials as any).data).forEach(([status, months]) => {
      const row: (string | number)[] = [status];
      let total = 0;
      (channelAnalysis.trials as any).months.forEach((m: number) => {
        const val = (months as any)[m] || 0;
        row.push(val);
        total += val;
      });
      row.push(total);
      trialsRows.push(row);
    });
    const trialsTotals: (string | number)[] = ['Totale complessivo'];
    let grandTotalTrials = 0;
    (channelAnalysis.trials as any).months.forEach((m: number) => {
      const monthTotal = Object.values((channelAnalysis.trials as any).data).reduce((sum, months) => 
        (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
      trialsTotals.push(monthTotal as number);
      grandTotalTrials += monthTotal as number;
    });
    trialsTotals.push(grandTotalTrials);
    trialsRows.push(trialsTotals);
    const nextRowProve = applyTableStyles(nextRowPresentati, 1, 'Prove', trialsHeaders, trialsRows, true);

    // 6. Fatturato (under Prove)
    const salesHeaders = ['Audioprotesista', ...(channelAnalysis.sales as any).months.map((m: number) => MONTH_NAMES[m - 1]), 'Totale'];
    const salesRows: any[][] = [];
    Object.entries((channelAnalysis.sales as any).data).forEach(([person, months]) => {
      const row: (string | number)[] = [person];
      let total = 0;
      (channelAnalysis.sales as any).months.forEach((m: number) => {
        const val = (months as any)[m] || 0;
        row.push(val);
        total += val;
      });
      row.push(total);
      salesRows.push(row);
    });
    const salesTotals: (string | number)[] = ['Totale complessivo'];
    let grandTotalSales = 0;
    (channelAnalysis.sales as any).months.forEach((m: number) => {
      const monthTotal = Object.values((channelAnalysis.sales as any).data).reduce((sum, months) => 
        (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
      salesTotals.push(monthTotal as number);
      grandTotalSales += monthTotal as number;
    });
    salesTotals.push(grandTotalSales);
    salesRows.push(salesTotals);

    applyTableStyles(nextRowProve, 1, 'Fatturato', salesHeaders, salesRows, true);

    // Adjust column widths
    sheet.columns.forEach(col => {
      col.width = 20;
    });

    // Write and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Analisi_Canali_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const availableBranches = useMemo(() => {
    const EXCLUDED_BRANCHES = [
      'alba', 'cuneo', 'castel del piano', 'mondovì', 
      'orbetello', 'grosseto', 'saluzzo', 'savigliano'
    ];
    const branches = new Set<string>();
    patientsData.forEach(p => {
      if (p.store) {
        const storeName = p.store.trim();
        if (!EXCLUDED_BRANCHES.includes(storeName.toLowerCase())) {
          branches.add(storeName);
        }
      }
    });
    return Array.from(branches).sort();
  }, [patientsData]);

  const availablePatientTypes = useMemo(() => {
    const types = new Set<string>();
    patientsData.forEach(p => {
      if (p.tipo) {
        types.add(p.tipo.trim());
      }
    });
    return Array.from(types).sort();
  }, [patientsData]);

  const availableCaps = useMemo(() => {
    const caps = new Set<string>();
    patientsData.forEach(p => {
      if (p.cap) {
        const c = String(p.cap).trim();
        if (c && c !== 'N/D') caps.add(c);
      }
    });
    return Array.from(caps).sort();
  }, [patientsData]);

  // Find all unique secondary channels and their Deposito di origine from salesData
  const secondaryToDeposito = useMemo(() => {
    const isRecapitoData = (s: any) => {
      const pChannel = String(s.canale || '').toLowerCase().trim();
      return pChannel === 'screening' || pChannel === 'promoter';
    };

    const map = new Map<string, string>();
    salesData.filter(isRecapitoData).forEach(s => {
      const sec = String(s.canaleSecondario || '').toLowerCase().trim();
      if (sec && sec !== 'n/d' && !map.has(sec)) {
        map.set(sec, normalizeDeposito(s.depositoOrigine));
      }
    });
    return map;
  }, [salesData]);

  // Derive list of unique normalized depositi for toggles
  const uniqueNormalizedDepositi = useMemo(() => {
    const deps = new Set<string>();
    secondaryToDeposito.forEach(dep => deps.add(dep));
    return Array.from(deps).sort();
  }, [secondaryToDeposito]);

  // Initialize visibility if empty
  useEffect(() => {
    if (visibleDepositi.length === 0 && uniqueNormalizedDepositi.length > 0) {
      const excluded = ['avezzano', 'cassino', 'viareggio', "l'aquila"];
      const initial = uniqueNormalizedDepositi.filter(dep => !excluded.includes(dep.toLowerCase().trim()));
      setVisibleDepositi(initial);
    }
  }, [uniqueNormalizedDepositi]);

  const comparison3YearData = useMemo(() => {
    if (salesData.length === 0) return { audioprotesisti: [], recapiti: [], availableDepositi: [] };

    const years = [comparisonAnchorYear, comparisonAnchorYear - 1, comparisonAnchorYear - 2];
    
    // Helper to calculate net stats for a subset of categories
    const getStats = (items: string[], type: 'audioprotesista' | 'canaleSecondario', filterFn?: (s: any) => boolean) => {
      return items.map(itemName => {
        const yearStats = years.map(year => {
          const monthly = new Array(12).fill(0);
          
          salesData.forEach(s => {
            if (s.year === year && String(s[type] || '').toLowerCase().trim() === itemName.toLowerCase().trim()) {
              if (!filterFn || filterFn(s)) {
                monthly[s.month] += (s.valore || 0);
              }
            }
          });
          
          creditNotes.forEach(c => {
            if (c.year === year && String(c[type] || '').toLowerCase().trim() === itemName.toLowerCase().trim()) {
              if (!filterFn || filterFn(c)) {
                monthly[c.month] -= (c.valore || 0);
              }
            }
          });
          
          return {
            year,
            monthly,
            total: monthly.reduce((a, b) => a + b, 0)
          };
        });
        
        return {
          name: itemName,
          years: yearStats
        };
      });
    };

    // Audioprotesisti list
    const audioprotesisti = Array.from(new Set(salesData.map(s => s.audioprotesista))).filter(n => n && n !== 'N/D').sort() as string[];
    
    // Recapiti logic: filter sales by primary channel "screening" or "promoter"
    const isRecapitoData = (s: any) => {
      const pChannel = String(s.canale || '').toLowerCase().trim();
      return pChannel === 'screening' || pChannel === 'promoter';
    };

    const recapitiInSales = Array.from(secondaryToDeposito.keys()) as string[];
    const allRecapiti = recapitiInSales.sort();

    return {
      audioprotesisti: getStats(audioprotesisti, 'audioprotesista'),
      recapiti: getStats(allRecapiti, 'canaleSecondario', isRecapitoData).map(stat => {
        const normalized = stat.name.toLowerCase().trim();
        return {
          ...stat,
          deposito: secondaryToDeposito.get(normalized) || 'N/D'
        };
      }).filter(stat => visibleDepositi.includes(normalizeDeposito(stat.deposito)))
      .sort((a, b) => {
        const depComp = normalizeDeposito(a.deposito).localeCompare(normalizeDeposito(b.deposito));
        if (depComp !== 0) return depComp;
        return a.name.localeCompare(b.name);
      }),
      availableDepositi: uniqueNormalizedDepositi as string[]
    };
  }, [salesData, creditNotes, comparisonAnchorYear, visibleDepositi, secondaryToDeposito, uniqueNormalizedDepositi]);

  // Set initial year when data is first loaded
  useEffect(() => {
    if (availableYears.length > 0 && !availableYears.includes(selectedYear)) {
      setSelectedYear(availableYears[0]);
    }
  }, [availableYears]);

  useEffect(() => {
    if (availableApptYears.length > 0 && !availableApptYears.includes(selectedApptYear)) {
      setSelectedApptYear(availableApptYears[0]);
    }
    if (availableApptYears.length > 0 && !availableApptYears.includes(selectedScreeningYear)) {
      setSelectedScreeningYear(availableApptYears[0]);
    }
  }, [availableApptYears]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (showStoreDropdown && !(event.target as HTMLElement).closest('.relative')) {
        setShowStoreDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [showStoreDropdown]);

  // Load stores from Firestore on mount
  useEffect(() => {
    const loadStoresAndCaps = async () => {
      setLoading(prev => ({ ...prev, stores: true }));
      try {
        const storesSnapshot = await getDocs(collection(db, 'stores'));
        const stores = storesSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        setStoresData(stores);

        const capsSnapshot = await getDocs(collection(db, 'storeCaps'));
        const caps = capsSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        setStoreCaps(caps);
      } catch (err) {
        console.error("Errore nel caricamento dei dati:", err);
      } finally {
        setLoading(prev => ({ ...prev, stores: false }));
      }
    };
    loadStoresAndCaps();
  }, []);

  // Click-away listener for dropdowns
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as HTMLElement;
      if (!target.closest('.relative')) {
        setShowStoreDropdown(false);
        setShowCapFilterDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleAssociateCap = async () => {
    if (!capInput.trim() || selectedStoreIds.length === 0) {
      setError("Inserisci un CAP e seleziona almeno una struttura.");
      return;
    }

    setLoading(prev => ({ ...prev, stores: true }));
    try {
      const batch = writeBatch(db);
      const newAssociations: any[] = [];
      const now = new Date().toISOString();

      selectedStoreIds.forEach(storeId => {
        // Check if association already exists
        const exists = storeCaps.some(sc => sc.storeId === storeId && sc.cap === capInput.trim());
        if (!exists) {
          const newDocRef = doc(collection(db, 'storeCaps'));
          const association = { 
            storeId, 
            cap: capInput.trim(),
            createdAt: now
          };
          batch.set(newDocRef, association);
          newAssociations.push({ id: newDocRef.id, ...association });
        }
      });

      if (newAssociations.length > 0) {
        await batch.commit();
        setStoreCaps(prev => [...prev, ...newAssociations]);
      }
      
      setCapInput('');
      setSelectedStoreIds([]);
      setError(null);
    } catch (err) {
      console.error("Errore nell'associazione CAP:", err);
      setError("Errore durante l'associazione dei CAP.");
    } finally {
      setLoading(prev => ({ ...prev, stores: false }));
    }
  };

  const toggleStoreSelection = (id: string) => {
    setSelectedStoreIds(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const filteredStores = useMemo(() => {
    let allStores = [...storesData];
    
    if (capSortConfig) {
      allStores.sort((a, b) => {
        let aValue = a[capSortConfig.key] || '';
        let bValue = b[capSortConfig.key] || '';
        
        if (capSortConfig.key === 'createdAt') {
          aValue = a.createdAt || '';
          bValue = b.createdAt || '';
        }

        if (aValue < bValue) return capSortConfig.direction === 'asc' ? -1 : 1;
        if (aValue > bValue) return capSortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    } else {
      allStores.sort((a, b) => (a.citta || '').localeCompare(b.citta || ''));
    }
    
    if (!capSearch.trim()) return allStores;
    const search = capSearch.toLowerCase();
    return allStores.filter(s => 
      s.nome.toLowerCase().includes(search) || 
      s.indirizzo.toLowerCase().includes(search) || 
      s.citta.toLowerCase().includes(search) ||
      (s.tipo || '').toLowerCase().includes(search) ||
      (s.deposito || '').toLowerCase().includes(search) ||
      (s.cap && s.cap.toLowerCase().includes(search))
    );
  }, [storesData, capSearch, capSortConfig]);

  const groupedStores = useMemo(() => {
    const grouped: Record<string, any[]> = {};
    filteredStores.forEach(s => {
      const dep = normalizeDeposito(s.deposito);
      if (!grouped[dep]) grouped[dep] = [];
      grouped[dep].push(s);
    });
    return Object.entries(grouped).sort((a, b) => a[0].localeCompare(b[0]));
  }, [filteredStores]);

  const processFile = async (file: File, type: 'sales' | 'credits' | 'trials' | 'patients' | 'calls' | 'appointments' | 'stores' | 'availability') => {
    setLoading(prev => ({ ...prev, [type]: true }));
    setError(null);
    
    try {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet) as any[];

        const EXCLUDE_LIST = ['askoltaora cuneo', 'askoltaora mondovì', 'askoltaora alba', 'daniele magnaldi'];

        const processed = json.map(row => {
          const findKey = (keys: string[]) => {
            const rowKeys = Object.keys(row);
            // 1. Exact match (case insensitive, trimmed)
            for (const k of keys) {
              const found = rowKeys.find(rk => rk.trim().toLowerCase() === k.toLowerCase());
              if (found) return found;
            }
            // 2. Starts with match
            for (const k of keys) {
              const found = rowKeys.find(rk => rk.trim().toLowerCase().startsWith(k.toLowerCase()));
              if (found) return found;
            }
            // 3. Includes match
            for (const k of keys) {
              const found = rowKeys.find(rk => rk.trim().toLowerCase().includes(k.toLowerCase()));
              if (found) return found;
            }
            return undefined;
          };

          const parseExcelDate = (val: any): { display: string, timestamp: number } => {
            if (!val) return { display: 'N/D', timestamp: 0 };
            let date: Date | null = null;
            
            if (val instanceof Date) {
              date = val;
            } else if (typeof val === 'number') {
              // Excel serial date
              date = new Date((val - 25569) * 86400 * 1000);
            } else if (typeof val === 'string') {
              const s = val.trim();
              // Handle DD/MM/YYYY or DD-MM-YYYY
              const ddmmyyyy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
              if (ddmmyyyy) {
                date = new Date(parseInt(ddmmyyyy[3]), parseInt(ddmmyyyy[2]) - 1, parseInt(ddmmyyyy[1]));
              } else {
                date = new Date(s);
              }
            }

            if (!date || isNaN(date.getTime())) {
              return { display: val ? String(val) : 'N/D', timestamp: 0 };
            }
            
            return { 
              display: date.toLocaleDateString('it-IT'), 
              timestamp: date.getTime() 
            };
          };

          const extractId = (val: any): string | undefined => {
            if (!val) return undefined;
            const s = String(val).trim();
            // Match ID:12345 or (ID: 12345) or [ID: 12345] or ID 12345
            const match = s.match(/(?:^|[^a-z0-9])ID[:\s]*([a-z0-9]+)/i);
            if (match) return match[1].trim();
            // If it's just a number, return it
            if (/^\d+$/.test(s)) return s;
            return undefined;
          };

          const extractName = (val: any): string => {
            if (!val) return 'N/D';
            let s = decodeHtml(String(val).trim());
            // Match "DA NAME (ID:ID)" or "NAME (ID:ID)"
            const match = s.match(/^(?:DA|A|PER)?\s*(.*?)\s*(?:[([ ]ID:|$)/i);
            if (match) return match[1].trim();
            // If no ID pattern, just clean up common prefixes
            return s.replace(/^(?:DA|A|PER)\s+/i, '').trim();
          };

          if (type === 'sales' || type === 'credits' || type === 'trials') {
            const nameKey = findKey(['Audioprotesista', 'Venditore', 'Socio', 'Intestatario', 'Nome']);
            const valueKey = findKey(['Valore to', 'Valore tot', 'Valore totale', 'Totale', 'Importo', 'Valore']);
            const dateKey = findKey(type === 'sales' ? ['Data documento'] : ['Data documento', 'Data doc', 'Data', 'Giorno', 'Data prova']);
            const customerKey = findKey(['Cliente', 'Nominativo', 'Ragione Sociale', 'Nome Cliente', 'Paziente']);
            const channelKey = findKey(['Canale primario', 'Canale', 'Marketing', 'Fonte', 'Provenienza', 'Campagna']);
            const secondaryChannelKey = findKey(['Canale secondario', 'Sottocanale', 'Sub-channel']);
            const depositoKey = findKey(['Deposito di origine', 'Deposito', 'Sede origine', 'Filiale origine']);
            const intermediarioKey = findKey(['Intermediario', 'Agente', 'Collaboratore']);
            const statoDocumentoKey = findKey(['Stato documento', 'Stato doc', 'Stato']);
            const quantitaKey = findKey(['Quantità totale', 'Qtà', 'Quantità', 'Qta']);

            if (!nameKey || !dateKey) return null;
            if ((type === 'sales' || type === 'credits') && !valueKey) return null;

            const rawName = String(row[nameKey]).trim();
            if (EXCLUDE_LIST.includes(rawName.toLowerCase())) return null;

            // Value parsing
            const parseCurrency = (val: any): number => {
              if (typeof val === 'number') {
                const s = String(val);
                const parts = s.split('.');
                // Se è un numero come 2.059 e ha 3 decimali, è quasi certamente un errore di lettura delle migliaia (Italiano)
                if (parts.length === 2 && parts[1].length === 3) {
                  return val * 1000;
                }
                return val;
              }
              if (!val) return 0;
              
              let s = String(val).trim().replace(/[^0-9,.-]/g, '');
              
              // Caso 1: Entrambi presenti (es. 1.315,29) -> Standard Italiano
              if (s.includes('.') && s.includes(',')) {
                return parseFloat(s.replace(/\./g, '').replace(',', '.')) || 0;
              }
              
              // Caso 2: Solo virgola (es. 1315,29) -> Decimale Italiano
              if (s.includes(',')) {
                return parseFloat(s.replace(',', '.')) || 0;
              }
              
              // Caso 3: Solo punto (es. 1315.29 o 2.059) -> Ambiguo
              if (s.includes('.')) {
                const parts = s.split('.');
                const lastPart = parts[parts.length - 1];
                
                // Se l'ultima parte ha 2 cifre, è un decimale (es. 1315.29)
                if (lastPart.length === 2) {
                  return parseFloat(s) || 0;
                }
                // Se l'ultima parte ha 3 cifre, è un separatore migliaia (es. 2.059)
                if (lastPart.length === 3) {
                  return parseFloat(s.replace(/\./g, '')) || 0;
                }
                // Default: prova a leggerlo come decimale
                return parseFloat(s) || 0;
              }
              
              return parseFloat(s) || 0;
            };

            // Date logic (prioritizing Data documento)
            let dateValue = row[dateKey];
            let date: Date;
            
            if (dateValue instanceof Date) {
              date = dateValue;
            } else if (typeof dateValue === 'number') {
              date = new Date((dateValue - 25569) * 86400 * 1000);
            } else {
              date = new Date(dateValue);
            }

            if (isNaN(date.getTime())) return null;

            const valore = valueKey ? parseCurrency(row[valueKey]) : 0;
            
            // Filter out sales < 600€ (only for ff.xlsx)
            if (type === 'sales' && valore < 600) return null;

            return {
              audioprotesista: rawName,
              cliente: customerKey ? String(row[customerKey]).trim() : 'N/D',
              canale: channelKey ? String(row[channelKey]).trim() : 'N/D',
              canaleSecondario: secondaryChannelKey ? String(row[secondaryChannelKey]).trim() : 'N/D',
              depositoOrigine: depositoKey ? normalizeDeposito(String(row[depositoKey])) : 'N/D',
              intermediario: intermediarioKey ? String(row[intermediarioKey]).trim() : 'N/D',
              valore: valore,
              statoDocumento: statoDocumentoKey ? String(row[statoDocumentoKey]).trim() : 'N/D',
              month: date.getMonth(),
              year: date.getFullYear(),
              fullDate: date.toLocaleDateString('it-IT'),
              quantita: quantitaKey ? Math.min(2, Math.max(0, parseInt(String(row[quantitaKey])) || 0)) : 0,
              rawColumn: valueKey
            };
          } else if (type === 'patients') {
            const idKey = findKey(['ID', 'Codice', 'Paziente ID', 'Cod. Paziente', 'Codice Paziente', 'ID Paziente']);
            const cognomeKey = findKey(['Cognome']);
            const nomeKey = findKey(['Nome']);
            const nominativoKey = findKey(['Nominativo', 'Paziente', 'Cliente', 'Nome Completo']);
            const phoneKey = findKey(['Telefono', 'Cellulare', 'Recapito', 'Tel', 'Cell']);
            const addressKey = findKey(['Indirizzo', 'Via', 'Residenza']);
            const storeKey = findKey(['Sede', 'Store', 'Filiale', 'Negozio', 'Sede di origine']);
            const cittaKey = findKey(['Città', 'Comune', 'Località']);
            const provinciaKey = findKey(['PR', 'Sigla', 'Provincia', 'Prov']);
            const capKey = findKey(['CAP', 'C.A.P.', 'Codice Postale']);
            const tipoKey = findKey(['Tipo', 'Stato Paziente', 'Categoria']);
            const channelKey = findKey(['Canale primario', 'Canale', 'Marketing', 'Fonte', 'Provenienza', 'Campagna']);
            const dateKey = findKey(['Creato il', 'Data creazione', 'Data', 'Creato', 'Giorno', 'Data inserimento']);

            // Collect all note-like fields
            const notesValues: string[] = [];
            ['Note', 'Osservazioni', 'Descrizione', 'Commenti'].forEach(k => {
              const found = findKey([k]);
              if (found && row[found]) {
                const val = decodeHtml(String(row[found]).trim());
                if (val && !notesValues.includes(val)) notesValues.push(val);
              }
            });
            const combinedNotes = notesValues.join(' | ');

            let fullName = 'N/D';
            let cognome = 'N/D';
            let nome = 'N/D';

            if (cognomeKey && nomeKey) {
              cognome = decodeHtml(String(row[cognomeKey]).trim());
              nome = decodeHtml(String(row[nomeKey]).trim());
              fullName = `${cognome} ${nome}`.trim();
            } else if (nominativoKey) {
              fullName = decodeHtml(String(row[nominativoKey]).trim());
              // Try to split nominativo into cognome and nome if possible
              const parts = fullName.split(' ');
              if (parts.length >= 2) {
                cognome = parts[0];
                nome = parts.slice(1).join(' ');
              } else {
                cognome = fullName;
                nome = '';
              }
            } else if (cognomeKey) {
              cognome = decodeHtml(String(row[cognomeKey]).trim());
              fullName = cognome;
            } else if (nomeKey) {
              nome = decodeHtml(String(row[nomeKey]).trim());
              fullName = nome;
            }

            const dateInfo = parseExcelDate(row[dateKey]);
            const date = dateInfo.timestamp ? new Date(dateInfo.timestamp) : new Date();

            const rawProvincia = provinciaKey ? String(row[provinciaKey]).trim() : '';
            // Stricter check for Province: must only contain letters and be reasonably short.
            // Exclude anything that looks like a number, a date, or an ID.
            let isProvinciaValid = true;
            if (!rawProvincia || rawProvincia === 'N/D' || rawProvincia === 'n/d') {
              isProvinciaValid = false;
            } else if (!/^[A-Za-zÀ-ÖØ-öø-ÿ\s]+$/.test(rawProvincia)) {
              // Contains non-alpha characters (numbers, symbols, etc.)
              isProvinciaValid = false;
            } else if (rawProvincia.length > 25) {
              // Too long for a province name/sigla
              isProvinciaValid = false;
            }

            const processedProvincia = isProvinciaValid ? rawProvincia : '';

            return {
              id: idKey ? String(row[idKey]).trim() : undefined,
              nomeCompleto: fullName,
              cognome: cognome,
              nome: nome,
              indirizzo: addressKey ? decodeHtml(String(row[addressKey]).trim()) : 'N/D',
              citta: cittaKey ? decodeHtml(String(row[cittaKey]).trim()) : 'N/D',
              provincia: processedProvincia,
              cap: capKey ? String(row[capKey]).trim() : 'N/D',
              note: combinedNotes,
              telefono: phoneKey ? String(row[phoneKey]).trim() : '',
              store: storeKey ? decodeHtml(String(row[storeKey]).trim()) : 'N/D',
              tipo: tipoKey ? decodeHtml(String(row[tipoKey]).trim()) : '',
              canale: channelKey ? String(row[channelKey]).trim() : 'N/D',
              month: date.getMonth(),
              year: date.getFullYear(),
              fullDate: date.toLocaleDateString('it-IT')
            };
          } else if (type === 'calls') {
            const idKey = findKey(['Contatti', 'Paziente ID', 'ID', 'Codice', 'Paziente', 'Contatto']);
            const nameKey = findKey(['Paziente', 'Nome', 'Nominativo', 'Contatti', 'Cliente', 'Contatto']);
            const dateKey = findKey(['Chiamato/a il', 'Chiamato il', 'Data telefonata', 'Data chiamata', 'Data', 'Giorno', 'Inizio', 'Ora', 'Creato']);
            const channelKey = findKey(['Canale primario', 'Canale', 'Marketing', 'Fonte', 'Provenienza', 'Campagna']);
            const contactKey = findKey(['Esito chiamata', 'Esito Chiamata', 'Esito telefonata', 'Risultato chiamata', 'Contatto', 'Esito', 'Risultato', 'Stato']);
            const audioprotesistaKey = findKey(['Impiegato', 'Audioprotesista', 'Operatore', 'Venditore', 'Audioprotesiste', 'Gestito da', 'Tecnico', 'In carico a', 'Professionista', 'Dipendente']);

            // Collect all note-like fields including "Esito chiamata"
            const notesValues: string[] = [];
            ['Esito chiamata', 'Esito', 'Note', 'Descrizione', 'Osservazioni', 'Commenti'].forEach(k => {
              const found = findKey([k]);
              if (found && row[found]) {
                const val = decodeHtml(String(row[found]).trim());
                if (val && !notesValues.includes(val)) notesValues.push(val);
              }
            });
            const combinedNotes = notesValues.join(' | ');

            const dateInfo = parseExcelDate(row[dateKey]);
            const date = dateInfo.timestamp ? new Date(dateInfo.timestamp) : new Date();

            return {
              pazienteId: idKey ? extractId(row[idKey]) : undefined,
              nome: nameKey ? extractName(row[nameKey]) : 'N/D',
              note: combinedNotes,
              data: dateInfo.display,
              timestamp: dateInfo.timestamp,
              canale: channelKey ? String(row[channelKey]).trim() : 'N/D',
              contatto: contactKey ? String(row[contactKey]).trim() : 'N/D',
              audioprotesista: audioprotesistaKey ? String(row[audioprotesistaKey]).trim() : 'N/D',
              month: date.getMonth(),
              year: date.getFullYear()
            };
          } else if (type === 'appointments') {
            const idKey = findKey(['Contatto', 'Paziente ID', 'ID', 'Codice', 'Paziente', 'Contatti']);
            const nameKey = findKey(['Paziente', 'Nome', 'Nominativo', 'Contatto', 'Cliente', 'Contatti']);
            const dateKey = findKey(['Previsto/a dalle ore', 'Data', 'Giorno', 'Inizio', 'Ora', 'Creato', 'Data appuntamento', 'Data visita']);
            const resultKey = findKey(['models/appointments.fields.appointment_result', 'Esito', 'Risultato', 'Stato', 'Contatto']);
            const channelKey = findKey(['models/appointments.fields.primary_channel_id', 'Canale primario', 'Canale', 'Marketing', 'Fonte', 'Provenienza', 'Campagna']);
            const tipoKey = findKey(['Tipo', 'Stato Paziente', 'Categoria']);
            const contactTypeKey = findKey(['models/appointments.fields.contacttype', 'Tipo contatto', 'Categoria contatto', 'Tipo Contatto', 'Tipo']);
            const audioprotesistaKey = findKey(['Audioprotesista', 'Operatore', 'Professionista', 'Gestito da', 'Tecnico', 'Medico']);
            const storeKey = findKey(['Sede', 'Sede di origine', 'Store', 'Negozio', 'Filiale', 'Punto vendita']);

            // Collect all note-like fields
            const notesValues: string[] = [];
            ['Note', 'Motivo', 'Descrizione', 'Osservazioni', 'Commenti'].forEach(k => {
              const found = findKey([k]);
              if (found && row[found]) {
                const val = decodeHtml(String(row[found]).trim());
                if (val && !notesValues.includes(val)) notesValues.push(val);
              }
            });
            const combinedNotes = notesValues.join(' | ');

            const esito = resultKey ? String(row[resultKey]).trim() : '';

            const dateInfo = parseExcelDate(row[dateKey]);
            const date = dateInfo.timestamp ? new Date(dateInfo.timestamp) : new Date();

            return {
              pazienteId: idKey ? extractId(row[idKey]) : undefined,
              nome: nameKey ? extractName(row[nameKey]) : 'N/D',
              note: combinedNotes,
              esito: esito,
              contatto: esito, // Alias for consistency
              data: dateInfo.display,
              timestamp: dateInfo.timestamp,
              canale: channelKey ? String(row[channelKey]).trim() : 'N/D',
              tipo: tipoKey ? String(row[tipoKey]).trim() : 'N/D',
              contactType: contactTypeKey ? String(row[contactTypeKey]).trim() : 'N/D',
              audioprotesista: audioprotesistaKey ? String(row[audioprotesistaKey]).trim() : 'N/D',
              sede: storeKey ? String(row[storeKey]).trim() : 'N/D',
              month: date.getMonth(),
              year: date.getFullYear()
            };
          } else if (type === 'stores') {
            const nameKey = findKey(['Nome', 'Struttura', 'Sede', 'Denominazione']);
            const addressKey = findKey(['Indirizzo', 'Via']);
            const cityKey = findKey(['Città', 'Comune', 'Località']);
            const capKey = findKey(['CAP', 'Codice Postale', 'Zip']);
            const phoneKey = findKey(['Telefono', 'Cellulare', 'Recapito']);
            const tipoKey = findKey(['Tipo', 'Tipologia', 'Categoria']);
            const notesKey = findKey(['Note', 'Osservazioni', 'Descrizione']);
            const depositoKey = findKey(['Deposito di origine', 'Deposito', 'Magazzino', 'Origine']);
            const createdAtKey = findKey(['Creato il', 'Data Creazione', 'Creato', 'Data']);

            if (!nameKey) return null;

            const storeData: any = {
              nome: decodeHtml(String(row[nameKey]).trim()),
              indirizzo: addressKey ? decodeHtml(String(row[addressKey]).trim()) : '',
              citta: cityKey ? decodeHtml(String(row[cityKey]).trim()) : '',
              cap: capKey ? String(row[capKey]).trim() : '',
              telefono: phoneKey ? String(row[phoneKey]).trim() : '',
              tipo: tipoKey ? decodeHtml(String(row[tipoKey]).trim()) : '',
              deposito: depositoKey ? normalizeDeposito(decodeHtml(String(row[depositoKey]))) : 'N/D',
              note: notesKey ? decodeHtml(String(row[notesKey]).trim()) : ''
            };

            if (createdAtKey) {
              const dateInfo = parseExcelDate(row[createdAtKey]);
              if (dateInfo.timestamp) {
                storeData.createdAt = new Date(dateInfo.timestamp).toISOString();
              }
            }

            return storeData;
          } else if (type === 'availability') {
            // Prioritize specific columns requested by the user: "Store" and "Previsto/a dalle"
            // We avoid generic names like "Data" or "Giorno" which might match creation dates or calendar templates
            const storeKey = findKey(['Store', 'Struttura', 'Sede', 'Nome', 'Recapito']);
            const dateKey = findKey(['Previsto/a alle', 'Previsto/a dalle ore', 'Previsto/a dalle', 'Data evento', 'Inizio']);
            const audioKey = findKey(['Audioprotesiste', 'Audioprotesista', 'Operatore']);
            
            if (!storeKey || !dateKey) return null;
            
            const rawStoreName = row[storeKey];
            const storeName = decodeHtml(String(rawStoreName || '').trim());
            
            // Skip rows with empty or generic store names
            if (!storeName || storeName.toLowerCase() === 'n/d' || storeName.length < 2) return null;
            
            const dateInfo = parseExcelDate(row[dateKey]);
            // If we can't parse a valid date, skip the row
            if (dateInfo.timestamp === 0) return null;

            const date = new Date(dateInfo.timestamp);

            return {
              storeName,
              audioprotesista: audioKey ? String(row[audioKey]).trim() : 'N/D',
              timestamp: dateInfo.timestamp,
              date: dateInfo.display,
              month: date.getMonth(),
              year: date.getFullYear()
            };
          }
          return null;
        }).filter(r => r !== null) as any[];

        if (type === 'stores') {
          const saveStores = async () => {
            try {
              const batch = writeBatch(db);
              const snapshot = await getDocs(collection(db, 'stores'));
              const existingStores = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() as any }));
              
              // Map existing stores by name for matching
              const nameToStore = new Map<string, any>();
              existingStores.forEach(s => nameToStore.set(s.nome.toLowerCase().trim(), s));

              const newStoreList: any[] = [];
              const processedNames = new Set<string>();

              processed.forEach(store => {
                const normalizedName = store.nome.toLowerCase().trim();
                if (processedNames.has(normalizedName)) return; // Avoid duplicates in the same file
                processedNames.add(normalizedName);

                const existing = nameToStore.get(normalizedName);
                if (existing) {
                  // Update existing store but keep ID
                  const storeRef = doc(db, 'stores', existing.id);
                  const updatedStore = { ...store };
                  
                  // Priority: 
                  // 1. Date from Excel file (store.createdAt)
                  // 2. Existing date in DB (existing.createdAt)
                  // 3. Current date (now)
                  if (!updatedStore.createdAt) {
                    if (existing.createdAt) {
                      updatedStore.createdAt = existing.createdAt;
                    } else {
                      updatedStore.createdAt = new Date().toISOString();
                    }
                  }
                  
                  batch.update(storeRef, updatedStore);
                  newStoreList.push({ id: existing.id, ...updatedStore });
                } else {
                  // Create new store
                  const newDoc = doc(collection(db, 'stores'));
                  const newStore = { 
                    ...store, 
                    createdAt: store.createdAt || new Date().toISOString() 
                  };
                  batch.set(newDoc, newStore);
                  newStoreList.push({ id: newDoc.id, ...newStore });
                }
              });

              // Optional: Delete stores that are no longer in the file? 
              // The user said "non cancellare", but usually a sync implies removing old ones.
              // However, to be safe and respect the "don't delete" sentiment, 
              // let's only delete if they are NOT in the new file AND have no associations?
              // Actually, let's just NOT delete anything for now to be 100% safe as requested.
              
              await batch.commit();
              
              // Reload to get the most up-to-date state
              const finalSnapshot = await getDocs(collection(db, 'stores'));
              const finalStores = finalSnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
              setStoresData(finalStores);
              
              // We DON'T setStoreCaps([]) anymore to preserve them
            } catch (err) {
              console.error(err);
              setError("Errore nel salvataggio delle strutture sul database");
            }
          };
          saveStores();
        } else if (type === 'sales') setSalesData(processed);
        else if (type === 'credits') setCreditNotes(processed);
        else if (type === 'trials') setTrialsData(processed);
        else if (type === 'patients') {
          setPatientsData(processed);
          setAnalyzedPatients([]);
        }
        else if (type === 'calls') {
          setCallsData(processed);
          setAnalyzedPatients([]);
        }
        else if (type === 'appointments') {
          setAppointmentsData(processed);
          setAnalyzedPatients([]);
        } else if (type === 'availability') {
          setAvailabilityData(processed);
        }

        setLoading(prev => ({ ...prev, [type]: false }));
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error(err);
      setError(`Errore durante il caricamento del file ${type}`);
      setLoading(prev => ({ ...prev, [type]: false }));
    }
  };

  const handleMultipleSalesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    Array.from(files).forEach((file: File) => {
      const fileName = file.name.toLowerCase();
      if (fileName.includes('ff.xlsx')) {
        processFile(file, 'sales');
      } else if (fileName.includes('rr.xlsx')) {
        processFile(file, 'credits');
      } else if (fileName.includes('pp.xlsx')) {
        processFile(file, 'trials');
      }
    });
  };

  const handleAudioDrillDown = (type: 'events' | 'calls' | 'appointments' | 'trials' | 'sales' | 'credits' | 'balance', monthIdx: number, filterValue?: string) => {
    const isFullYear = monthIdx === -1;
    const title = selectedAudioPro;
    let subtitle = `${isFullYear ? 'Bilancio Annuale' : MONTHS[monthIdx]} ${selectedAudioYear}`;
    let data: any[] = [];
    let drillDownType: 'patients' | 'calls' | 'appointments' | 'trials' | 'sales' | 'availability' = 'sales';

    const baseFilter = (d: any) => 
      String(d.audioprotesista || '').toLowerCase().trim() === selectedAudioPro.toLowerCase().trim() && 
      d.year === selectedAudioYear && 
      (isFullYear || d.month === monthIdx);

    if (type === 'events') {
      drillDownType = 'availability';
      data = availabilityData.filter(d => 
        baseFilter(d) && 
        (!filterValue || d.storeName === filterValue) &&
        !['domicilio - fabriano', 'fabriano', 'domicilio - genova', 'domicilio', 'genova', 'domicilio - roma pallavicini', 'roma - pallavicini', 'aprilia', 'ricerca struttura'].includes(String(d.storeName || '').toLowerCase().trim())
      );
      subtitle += ` • Store: ${filterValue || 'Tutti'}`;
    } else if (type === 'calls') {
      drillDownType = 'calls';
      data = callsData.filter(d => baseFilter(d) && (!filterValue || d.contatto === filterValue));
      subtitle += ` • Esito: ${filterValue || 'Tutti'}`;
    } else if (type === 'appointments') {
      drillDownType = 'appointments';
      data = appointmentsData.filter(d => {
        const dType = String(d.tipo || '').toLowerCase().trim();
        const dContactType = String(d.contactType || '').toLowerCase().trim();
        const isCorrectType = dType.includes('controllo udito') || dType.includes('prima visita') ||
                              dContactType.includes('controllo udito') || dContactType.includes('prima visita');
        
        const esito = String(d.esito || '').toLowerCase().trim();
        const isPresentato = esito === 'si è presentato' || esito === 'si e presentato';

        return baseFilter(d) && 
          (!filterValue || d.contactType === filterValue || d.tipo === filterValue) &&
          isPresentato &&
          isCorrectType;
      });
      subtitle += ` • Tipo: ${filterValue || 'Controllo Udito / Prima Visita'}`;
    } else if (type === 'trials') {
      drillDownType = 'trials';
      data = trialsData.filter(d => 
        baseFilter(d) && 
        String(d.statoDocumento || '').toLowerCase().trim() === 'prova aperta'
      );
      subtitle += ` • Prove Aperte`;
    } else if (type === 'sales') {
      drillDownType = 'sales';
      data = salesData.filter(d => baseFilter(d));
      subtitle += ` • Vendite`;
    } else if (type === 'credits') {
      drillDownType = 'sales'; // We use sales UI for credits too, they share fields
      data = creditNotes.filter(d => baseFilter(d));
      subtitle += ` • Storni`;
    } else if (type === 'balance') {
      drillDownType = 'sales';
      const sales = salesData.filter(d => baseFilter(d));
      const credits = creditNotes.filter(d => baseFilter(d));
      data = [...sales, ...credits];
      subtitle += ` • Saldo (Vendite + Storni)`;
    }

    if (data.length > 0) {
      setChannelDrillDown({
        title,
        subtitle,
        data: data.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0)),
        type: drillDownType
      });
    }
  };

  const handleAppointmentsDrillDown = (audioprotesista: string | null, monthIdx: number, subGroupValue?: string) => {
    const isFullYear = monthIdx === -1;
    const title = audioprotesista || 'Tutti gli Audioprotesisti';
    let subtitle = `${isFullYear ? 'Totale' : MONTHS[monthIdx]} ${selectedApptYear}`;
    
    if (subGroupValue) {
      subtitle += ` • ${appointmentsAnalysisView === 'canali' ? 'Canale' : 'Tipo'}: ${subGroupValue}`;
    }

    const data = appointmentsData.filter(a => {
      if (a.year !== selectedApptYear) return false;
      if (!isFullYear && a.month !== monthIdx) return false;
      
      const esito = String(a.esito || '').toLowerCase().trim();
      const tipo = String(a.tipo || '').toLowerCase().trim();
      
      const matchEsito = esito === 'si è presentato' || esito === 'si e presentato' || esito === 'presentato';
      const matchTipo = tipo === 'controllo udito' || tipo === 'prima visita';
      
      if (!matchEsito || !matchTipo) return false;

      if (audioprotesista) {
        const pro = a.audioprotesista || 'Da assegnare';
        if (pro !== audioprotesista) return false;
      }

      if (subGroupValue) {
        const itemSubGroup = appointmentsAnalysisView === 'canali' ? (a.canale || 'N/D') : (a.contactType || 'N/D');
        if (itemSubGroup !== subGroupValue) return false;
      }

      return true;
    });

    if (data.length > 0) {
      setChannelDrillDown({
        title,
        subtitle,
        data: data.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0)),
        type: 'appointments'
      });
    }
  };

  const handleScreeningDrillDown = (audioprotesista: string | null, monthIdx: number, storeValue?: string) => {
    const isFullYear = monthIdx === -1;
    const title = audioprotesista || 'Tutti gli Audioprotesisti';
    let subtitle = `${isFullYear ? 'Totale' : MONTHS[monthIdx]} ${selectedScreeningYear}`;
    
    if (storeValue) {
      subtitle += ` • Store: ${storeValue}`;
    }

    const data = appointmentsData.filter(a => {
      if (a.year !== selectedScreeningYear) return false;
      if (!isFullYear && a.month !== monthIdx) return false;
      
      const esito = String(a.esito || '').toLowerCase().trim();
      const tipo = String(a.tipo || '').toLowerCase().trim();
      
      const matchEsito = esito === 'si è presentato' || esito === 'si e presentato' || esito === 'presentato';
      const matchTipo = tipo === 'controllo udito' || tipo === 'prima visita';
      
      if (!matchEsito || !matchTipo) return false;

      if (audioprotesista) {
        const pro = a.audioprotesista || 'Da assegnare';
        if (pro !== audioprotesista) return false;
      }

      if (storeValue) {
        const itemStore = a.sede || 'N/D';
        if (itemStore !== storeValue) return false;
      }

      return true;
    });

    if (data.length > 0) {
      setChannelDrillDown({
        title,
        subtitle,
        data: data.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0)),
        type: 'appointments'
      });
    }
  };

  const handleMultiplePatientsUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    Array.from(files).forEach((file: File) => {
      const fileName = file.name.toLowerCase();
      if (fileName.includes('ii.xlsx')) {
        processFile(file, 'patients');
      } else if (fileName.includes('tt.xlsx')) {
        processFile(file, 'calls');
      } else if (fileName.includes('aa.xlsx')) {
        processFile(file, 'appointments');
      } else if (fileName.includes('store.xlsx')) {
        processFile(file, 'stores');
      } else if (fileName.includes('dd.xlsx')) {
        processFile(file, 'availability');
      }
    });
  };

  const handleNoahUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setLoading(prev => ({ ...prev, noah: true }));
    setError(null);
    const allPatients: any[] = [];

    const processSingleFile = (file: File): Promise<any[]> => {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (event) => {
          try {
            const xmlData = event.target?.result as string;
            const parser = new XMLParser({
              ignoreAttributes: false,
              attributeNamePrefix: "@_",
              removeNSPrefix: true,
            });
            const jsonObj = parser.parse(xmlData);
            
            let patients: any[] = [];
            
            const findPatients = (obj: any) => {
              if (!obj || typeof obj !== 'object') return;
              
              if (Array.isArray(obj)) {
                obj.forEach(item => findPatients(item));
                return;
              }

              const keys = Object.keys(obj);
              const patientTags = ['Patient', 'Paziente', 'Client', 'PatientRecord', 'ClientRecord', 'PazienteRecord', 'Member'];
              
              for (const tag of patientTags) {
                // Match exact tag or tag with namespace (e.g. pt:Patient)
                const matchKey = keys.find(k => k === tag || k.endsWith(':' + tag));
                if (matchKey) {
                  const p = obj[matchKey];
                  if (Array.isArray(p)) {
                    patients.push(...p);
                  } else {
                    patients.push(p);
                  }
                  // Continue searching inside this object as well
                }
              }

              Object.values(obj).forEach(val => findPatients(val));
            };

            findPatients(jsonObj);

            if (patients.length === 0) {
              const findPotentialPatients = (obj: any) => {
                if (!obj || typeof obj !== 'object') return;
                if (Array.isArray(obj)) {
                  obj.forEach(item => {
                    if (item && (item.FirstName || item.LastName || item.Nome || item.Cognome || item.Name || item.FullName || item.GivenName || item.Surname)) {
                      patients.push(item);
                    } else {
                      findPotentialPatients(item);
                    }
                  });
                } else {
                  if (obj.FirstName || obj.LastName || obj.Nome || obj.Cognome || obj.Name || obj.FullName || obj.GivenName || obj.Surname) {
                    patients.push(obj);
                  } else {
                    Object.values(obj).forEach(val => findPotentialPatients(val));
                  }
                }
              };
              findPotentialPatients(jsonObj);
            }

            const filtered = patients.map(p => {
              const extractValue = (val: any): string => {
                if (val === null || val === undefined) return '';
                if (typeof val === 'object') {
                  // Handle fast-xml-parser's text node or just take the first property that looks like a value
                  const text = val['#text'] || val.text || val.Value || val.val || val.valore || val.dato || '';
                  if (text) return String(text).trim();
                  // Fallback: if it's an object with only one key, it might be the value
                  const values = Object.values(val).filter(v => typeof v !== 'object');
                  if (values.length > 0) return String(values[0]).trim();
                  return '';
                }
                return String(val).trim();
              };

              const findField = (obj: any, fieldNames: string[]): any => {
                if (!obj || typeof obj !== 'object') return undefined;
                
                const keys = Object.keys(obj);
                // Check current level
                for (const name of fieldNames) {
                  if (obj[name] !== undefined) return obj[name];
                  if (obj['@_' + name] !== undefined) return obj['@_' + name];
                  
                  // Check for namespace prefixed keys (e.g. pt:FirstName)
                  const nsKey = keys.find(k => k.endsWith(':' + name));
                  if (nsKey) return obj[nsKey];
                }

                // Check nested levels
                const queue = [...Object.values(obj)];
                let depth = 0;
                while (queue.length > 0 && depth < 200) {
                  const current = queue.shift();
                  if (current && typeof current === 'object') {
                    const currentKeys = Object.keys(current);
                    for (const name of fieldNames) {
                      if (current[name] !== undefined) return current[name];
                      if (current['@_' + name] !== undefined) return current['@_' + name];
                      
                      const nsKey = currentKeys.find(k => k.endsWith(':' + name));
                      if (nsKey) return current[nsKey];
                    }
                    if (Array.isArray(current)) {
                      queue.push(...current);
                    } else {
                      queue.push(...Object.values(current));
                    }
                  }
                  depth++;
                }
                return undefined;
              };

              const nomeRaw = findField(p, ['FirstName', 'Nome', 'Firstname', 'GivenName']);
              const cognomeRaw = findField(p, ['LastName', 'Cognome', 'Lastname', 'Surname']);
              const fullNameRaw = findField(p, ['Name', 'FullName', 'Nominativo']);
              
              const nome = extractValue(nomeRaw);
              const cognome = extractValue(cognomeRaw);
              const fullName = extractValue(fullNameRaw);
              
              const finalNome = nome || (fullName ? fullName.split(' ').slice(0, -1).join(' ') : '');
              const finalCognome = cognome || (fullName ? fullName.split(' ').slice(-1)[0] : '');

              const telefonoRaw = findField(p, ['MobilePhone', 'Phone', 'Telefono', 'Mobile', 'Cellulare', 'Telephone', 'HomePhone', 'WorkPhone', 'PhoneMobile', 'MobileNumber', 'PhoneNumber']);
              const telefono = extractValue(telefonoRaw);

              const mobileRaw = findField(p, ['MobilePhone', 'Mobile', 'Cellulare']);
              const mobile = extractValue(mobileRaw);
              
              const homeRaw = findField(p, ['HomePhone', 'Telephone', 'Phone', 'Telefono']);
              const home = extractValue(homeRaw);

              const createDateRaw = findField(p, ['CreateDate', 'CreatedDate', 'DateCreated', 'CreationDate']);
              const createDate = extractValue(createDateRaw);

              // Exclude if no phone, or if any found phone is "0"
              if (!telefono || telefono === '0' || mobile === '0' || home === '0') return null;

              let hasSpecificValue = false;
              
              let value1000: number | null = null;
              
              const checkAudiogram = (obj: any) => {
                if (!obj || typeof obj !== 'object') return;
                
                if (Array.isArray(obj)) {
                  obj.forEach(item => checkAudiogram(item));
                  return;
                }

                // Check for frequency and level in this object
                const freqTags = ['StimulusFrequency', 'Frequency', 'Freq', 'Hz', 'Frequenza'];
                const levelTags = ['StimulusLevel', 'Level', 'Intensity', 'HL', 'Valore', 'Value', 'Intensita', 'Livello'];

                let freq: any = undefined;
                let level: any = undefined;

                const objKeys = Object.keys(obj);
                for (const tag of freqTags) {
                  if (obj[tag] !== undefined) { freq = obj[tag]; break; }
                  if (obj['@_' + tag] !== undefined) { freq = obj['@_' + tag]; break; }
                  const nsKey = objKeys.find(k => k.endsWith(':' + tag));
                  if (nsKey) { freq = obj[nsKey]; break; }
                }

                for (const tag of levelTags) {
                  if (obj[tag] !== undefined) { level = obj[tag]; break; }
                  if (obj['@_' + tag] !== undefined) { level = obj['@_' + tag]; break; }
                  const nsKey = objKeys.find(k => k.endsWith(':' + tag));
                  if (nsKey) { level = obj[nsKey]; break; }
                }

                if (freq !== undefined && level !== undefined) {
                  const fStr = extractValue(freq).toLowerCase().replace(',', '.');
                  const lStr = extractValue(level).toLowerCase().replace(',', '.');
                  
                  let f = Number(fStr.replace(/[^0-9.]/g, ''));
                  if (fStr.includes('khz')) f *= 1000;
                  
                  const l = Number(lStr.replace(/[^0-9.-]/g, ''));
                  
                  // Use a small range for frequency to handle floating point or slight variations
                  if (Math.abs(f - 1000) < 1) {
                    if (value1000 === null || l > value1000) {
                      value1000 = l;
                    }
                    if (l > noahThreshold) {
                      hasSpecificValue = true;
                    }
                  }
                }

                // Recurse into children
                Object.entries(obj).forEach(([key, val]) => {
                  if (typeof val === 'object') {
                    // Skip UncomfortableLevel tags
                    if (key === 'UncomfortableLevel' || key.endsWith(':UncomfortableLevel')) {
                      return;
                    }
                    checkAudiogram(val);
                  }
                });
              };

              checkAudiogram(p);

              if (!hasSpecificValue) return null;

              return {
                nome: finalNome || 'N/D',
                cognome: finalCognome || 'N/D',
                telefono,
                valore1000: value1000,
                dataCreazione: createDate
              };
            }).filter(p => p !== null) as any[];

            resolve(filtered);
          } catch (err) {
            console.error("Error parsing XML:", err);
            reject(err);
          }
        };
        reader.onerror = () => reject(new Error("Errore durante la lettura del file."));
        reader.readAsText(file);
      });
    };

    try {
      const results = await Promise.all((Array.from(files) as File[]).map(file => processSingleFile(file)));
      const combined = results.flat();

      // Sort by creation date (oldest to newest)
      combined.sort((a, b) => {
        if (!a.dataCreazione) return 1;
        if (!b.dataCreazione) return -1;
        const dateA = new Date(a.dataCreazione).getTime();
        const dateB = new Date(b.dataCreazione).getTime();
        return dateA - dateB;
      });

      setNoahPatients(combined);
      console.log(`Noah Analysis: Found ${combined.length} patients matching criteria across ${files.length} files.`);
    } catch (err) {
      setError("Errore durante la lettura di uno o più file XML. Assicurati che siano file Noah validi.");
    } finally {
      setLoading(prev => ({ ...prev, noah: false }));
    }
  };

  const handleDownloadNoahExcel = async () => {
    if (noahPatients.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Analisi Noah');

    worksheet.columns = [
      { header: 'Cognome', key: 'cognome', width: 25 },
      { header: 'Nome', key: 'nome', width: 25 },
      { header: 'Telefono', key: 'telefono', width: 20 },
      { header: 'Valore 1000Hz (dB)', key: 'valore1000', width: 20 },
      { header: 'Data Creazione', key: 'dataCreazione', width: 25 }
    ];

    // Style header
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };

    noahPatients.forEach(p => {
      worksheet.addRow({
        cognome: p.cognome,
        nome: p.nome,
        telefono: p.telefono,
        valore1000: p.valore1000 !== null ? `${p.valore1000} dB` : 'N/D',
        dataCreazione: p.dataCreazione || 'N/D'
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Analisi_Noah_${new Date().toISOString().split('T')[0]}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  };

  const aggregatedData = useMemo(() => {
    // Structure: { Name: { sales: [12], credits: [12], channelMonthly: { [channelName]: [12] } } }
    const stats: Record<string, { 
      sales: number[], 
      credits: number[], 
      channelMonthly: Record<string, number[]> 
    }> = {};

    salesData.forEach(sale => {
      if (sale.year === selectedYear) {
        if (!stats[sale.audioprotesista]) {
          stats[sale.audioprotesista] = { 
            sales: new Array(12).fill(0), 
            credits: new Array(12).fill(0),
            channelMonthly: {}
          };
        }
        stats[sale.audioprotesista].sales[sale.month] += sale.valore;
        
        const channel = sale.canale || 'N/D';
        if (!stats[sale.audioprotesista].channelMonthly[channel]) {
          stats[sale.audioprotesista].channelMonthly[channel] = new Array(12).fill(0);
        }
        stats[sale.audioprotesista].channelMonthly[channel][sale.month] += sale.valore;
      }
    });

    creditNotes.forEach(note => {
      if (note.year === selectedYear) {
        if (!stats[note.audioprotesista]) {
          stats[note.audioprotesista] = { 
            sales: new Array(12).fill(0), 
            credits: new Array(12).fill(0),
            channelMonthly: {}
          };
        }
        stats[note.audioprotesista].credits[note.month] += note.valore;
        
        const channel = note.canale || 'N/D';
        if (!stats[note.audioprotesista].channelMonthly[channel]) {
          stats[note.audioprotesista].channelMonthly[channel] = new Array(12).fill(0);
        }
        stats[note.audioprotesista].channelMonthly[channel][note.month] -= note.valore;
      }
    });

    return Object.entries(stats).sort((a, b) => a[0].localeCompare(b[0]));
  }, [salesData, creditNotes, selectedYear]);

  const handleCellClick = (name: string | null, monthIdx: number, channel?: string, year?: number, type: 'audioprotesista' | 'recapiti' = 'audioprotesista', deposito?: string) => {
    const isFullYear = monthIdx === -1;
    const targetYear = year || selectedYear;

    const filterFn = (s: any) => {
      let matchesItem = false;
      if (type === 'audioprotesista') {
        matchesItem = !name || s.audioprotesista === name;
      } else {
        // For recapiti, if name is provided it's the structure (canaleSecondario)
        // If deposito is provided it's the group filter
        if (name) {
          const targetName = name.toLowerCase().trim();
          matchesItem = String(s.canaleSecondario || '').toLowerCase().trim() === targetName;
        } else if (deposito) {
          matchesItem = normalizeDeposito(s.depositoOrigine) === deposito;
        } else {
          matchesItem = true;
        }
        
        const pChannel = String(s.canale || '').toLowerCase().trim();
        matchesItem = matchesItem && (pChannel === 'screening' || pChannel === 'promoter');
      }

      return matchesItem && 
             (isFullYear || s.month === monthIdx) && 
             s.year === targetYear &&
             (!channel || s.canale === channel);
    };

    const sales = salesData.filter(filterFn);
    const credits = creditNotes.filter(filterFn);
    
    if (sales.length > 0 || credits.length > 0) {
      setDrillDown({
        name: name || (channel ? `Canale: ${channel}` : 'Totale'),
        month: monthIdx,
        channel,
        sales,
        credits: credits.map(c => ({ ...c, audioprotesista: c.audioprotesista || name })), // Ensure audioprotesista is present
        year: targetYear
      });
    }
  };

  const handleScreeningCellClick = (deposito: string | null, monthIdx: number, secondaryChannel?: string) => {
    const isFullYear = monthIdx === -1;
    const sales = salesData.filter(s => 
      s.year === selectedYear &&
      (s.canale?.toUpperCase() === 'SCREENING' || s.canale?.toUpperCase() === 'PROMOTER') &&
      (!deposito || s.depositoOrigine === deposito) &&
      (isFullYear || s.month === monthIdx) &&
      (!secondaryChannel || s.canaleSecondario === secondaryChannel)
    );
    const credits = creditNotes.filter(c => 
      c.year === selectedYear &&
      (c.canale?.toUpperCase() === 'SCREENING' || c.canale?.toUpperCase() === 'PROMOTER') &&
      (!deposito || c.depositoOrigine === deposito) &&
      (isFullYear || c.month === monthIdx) &&
      (!secondaryChannel || c.canaleSecondario === secondaryChannel)
    );
    
    if (sales.length > 0 || credits.length > 0) {
      setDrillDown({
        name: deposito || (secondaryChannel ? `Canale: ${secondaryChannel}` : 'Totale Screening'),
        month: monthIdx,
        channel: secondaryChannel,
        sales,
        credits
      });
    }
  };

  const handleMediciCellClick = (audioprotesista: string | null, monthIdx: number, secondaryChannel?: string) => {
    const isFullYear = monthIdx === -1;
    const sales = salesData.filter(s => 
      s.year === selectedYear &&
      s.canale?.toUpperCase() === 'MEDICO' &&
      (!audioprotesista || s.audioprotesista === audioprotesista) &&
      (isFullYear || s.month === monthIdx) &&
      (!secondaryChannel || s.canaleSecondario === secondaryChannel)
    );
    const credits = creditNotes.filter(c => 
      c.year === selectedYear &&
      c.canale?.toUpperCase() === 'MEDICO' &&
      (!audioprotesista || c.audioprotesista === audioprotesista) &&
      (isFullYear || c.month === monthIdx) &&
      (!secondaryChannel || c.canaleSecondario === secondaryChannel)
    );
    
    if (sales.length > 0 || credits.length > 0) {
      setDrillDown({
        name: audioprotesista || (secondaryChannel ? `Canale: ${secondaryChannel}` : 'Totale Medici'),
        month: monthIdx,
        channel: secondaryChannel,
        sales,
        credits
      });
    }
  };

  const handleDownloadComparisonExcel = async () => {
    if (comparison3YearData.audioprotesisti.length === 0 && comparison3YearData.recapiti.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Comparazioni 3 Anni');
    const years = [comparisonAnchorYear - 2, comparisonAnchorYear - 1, comparisonAnchorYear];

    const formatTable = (title: string, data: any[], startRow: number, isRecapiti = false) => {
      let currentRow = startRow;

      // Section Title
      const titleRow = sheet.getRow(currentRow);
      titleRow.getCell(1).value = title;
      titleRow.getCell(1).font = { bold: true, size: 14 };
      currentRow++;

      // Headers - First Level (Months)
      const headerRow1 = sheet.getRow(currentRow);
      headerRow1.getCell(1).value = 'Nome';
      headerRow1.getCell(1).font = { bold: true };
      headerRow1.getCell(1).border = { bottom: { style: 'thin' }, right: { style: 'medium' } };

      MONTHS.forEach((month, mIdx) => {
        const colStart = 2 + (mIdx * 3);
        const cell = headerRow1.getCell(colStart);
        cell.value = month;
        cell.alignment = { horizontal: 'center' };
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        cell.border = { 
          top: { style: 'thin' }, 
          left: { style: 'thin' }, 
          right: { style: 'medium' }, 
          bottom: { style: 'thin' } 
        };
        sheet.mergeCells(currentRow, colStart, currentRow, colStart + 2);
      });

      // Totale Header
      const totalColStart = 2 + (MONTHS.length * 3);
      const totalCell = headerRow1.getCell(totalColStart);
      totalCell.value = 'Totale YTD';
      totalCell.alignment = { horizontal: 'center' };
      totalCell.font = { bold: true };
      totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD0D0D0' } };
      totalCell.border = { 
        top: { style: 'thin' }, 
        left: { style: 'thin' }, 
        right: { style: 'medium' }, 
        bottom: { style: 'thin' } 
      };
      sheet.mergeCells(currentRow, totalColStart, currentRow, totalColStart + 2);
      currentRow++;

      // Headers - Second Level (Years)
      const headerRow2 = sheet.getRow(currentRow);
      MONTHS.forEach((_, mIdx) => {
        years.forEach((year, yIdx) => {
          const col = 2 + (mIdx * 3) + yIdx;
          const cell = headerRow2.getCell(col);
          cell.value = year;
          cell.alignment = { horizontal: 'center' };
          cell.font = { bold: true };
          cell.border = { 
            left: { style: 'thin' }, 
            right: { style: yIdx === 2 ? 'medium' : 'thin' }, 
            bottom: { style: 'medium' } 
          };
        });
      });

      // Totale Years
      years.forEach((year, yIdx) => {
        const col = totalColStart + yIdx;
        const cell = headerRow2.getCell(col);
        cell.value = year;
        cell.alignment = { horizontal: 'center' };
        cell.font = { bold: true };
        cell.border = { 
          left: { style: 'thin' }, 
          right: { style: yIdx === 2 ? 'medium' : 'thin' }, 
          bottom: { style: 'medium' } 
        };
      });
      currentRow++;

      const renderRows = (items: any[]) => {
        items.forEach((row) => {
          const dr = sheet.getRow(currentRow);
          dr.getCell(1).value = row.name;
          dr.getCell(1).border = { right: { style: 'medium' }, bottom: { style: 'thin' } };

          MONTHS.forEach((_, mIdx) => {
            years.forEach((_, yIdx) => {
              const col = 2 + (mIdx * 3) + yIdx;
              const dataIdx = 2 - yIdx;
              const val = row.years[dataIdx].monthly[mIdx];
              const cell = dr.getCell(col);
              cell.value = Math.round(val);
              cell.numFmt = '#,##0';
              cell.alignment = { horizontal: 'right' };
              cell.border = { 
                right: { style: yIdx === 2 ? 'medium' : 'thin' }, 
                bottom: { style: 'thin' } 
              };
            });
          });

          // Total Values
          years.forEach((_, yIdx) => {
            const col = totalColStart + yIdx;
            const dataIdx = 2 - yIdx;
            const cell = dr.getCell(col);
            cell.value = Math.round(row.years[dataIdx].total);
            cell.numFmt = '#,##0';
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'right' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: yIdx === 2 ? 'FFFFFF00' : 'FFF0F0F0' } };
            cell.border = { 
              right: { style: yIdx === 2 ? 'medium' : 'thin' }, 
              bottom: { style: 'thin' } 
            };
          });
          currentRow++;
        });
      };

      if (isRecapiti) {
        const groups: Record<string, any[]> = {};
        data.forEach(row => {
          const dep = normalizeDeposito(row.deposito);
          if (!groups[dep]) groups[dep] = [];
          groups[dep].push(row);
        });
        
        Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0])).forEach(([deposito, items]) => {
          // Calculate group totals
          const groupTotals = years.map((_, yIdx) => {
            const monthly = new Array(12).fill(0);
            items.forEach(it => {
              // it.years is [2026_stats, 2025_stats, 2024_stats]
              // Excel years are [2024, 2025, 2026]
              // so Excel yIdx 0 (2024) -> it.years[2]
              const dataIdxForTotals = 2 - yIdx;
              it.years[dataIdxForTotals].monthly.forEach((val: number, mIdx: number) => {
                monthly[mIdx] += val;
              });
            });
            return {
              monthly,
              total: monthly.reduce((a: number, b: number) => a + b, 0)
            };
          });

          // Group Header
          const groupRow = sheet.getRow(currentRow);
          groupRow.getCell(1).value = `DEPOSITO: ${deposito}`;
          groupRow.getCell(1).font = { bold: true };
          groupRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
          
          MONTHS.forEach((_, mIdx) => {
            years.forEach((_, yIdx) => {
              const col = 2 + (mIdx * 3) + yIdx;
              const cell = groupRow.getCell(col);
              cell.value = Math.round(groupTotals[yIdx].monthly[mIdx]);
              cell.numFmt = '#,##0';
              cell.font = { bold: true };
              cell.alignment = { horizontal: 'right' };
              cell.border = { 
                right: { style: yIdx === 2 ? 'medium' : 'thin' }, 
                bottom: { style: 'thin' } 
              };
            });
          });

          // Group Totals YTD
          years.forEach((_, yIdx) => {
            const col = totalColStart + yIdx;
            const cell = groupRow.getCell(col);
            cell.value = Math.round(groupTotals[yIdx].total);
            cell.numFmt = '#,##0';
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'right' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: yIdx === 2 ? 'FFFFFF00' : 'FFD0D0D0' } };
            cell.border = { 
              right: { style: yIdx === 2 ? 'medium' : 'thin' }, 
              bottom: { style: 'thin' } 
            };
          });

          currentRow++;
          renderRows(items);
        });
      } else {
        renderRows(data);
      }

      return currentRow + 2;
    };

    let nextStart = 2;
    nextStart = formatTable('ANALISI AUDIOPROTESISTI', comparison3YearData.audioprotesisti, nextStart);
    formatTable('ANALISI RECAPITI', comparison3YearData.recapiti, nextStart, true);

    // Set column widths
    sheet.getColumn(1).width = 30;
    for (let i = 2; i <= 2 + (MONTHS.length * 3) + 3; i++) {
      sheet.getColumn(i).width = 10;
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Comparazione_3Anni_${comparisonAnchorYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const totalSales = aggregatedData.reduce((acc, [_, data]) => acc + data.sales.reduce((mAcc, m) => mAcc + m, 0) - data.credits.reduce((mAcc, m) => mAcc + m, 0), 0);

  const handleInteressanti = () => {
    if (patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0) {
      setError("Carica tutti e tre i file (Anagrafica, Telefonate, Appuntamenti) prima di procedere.");
      return;
    }

    if (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim()) {
      setError("Seleziona una filiale, una struttura esterna o un CAP prima di procedere.");
      return;
    }

    setIsAnalyzing(true);
    setError(null);

    // Simulate a small delay for better UX
    setTimeout(() => {
      const now = new Date();
      const threeMonthsAgo = now.getTime() - (90 * 24 * 60 * 60 * 1000);
      const twoMonthsAgo = now.getTime() - (60 * 24 * 60 * 60 * 1000);
      const filteredPatients = patientsData.filter(p => {
        const isCorrectType = !['Deceduto', 'Cliente', 'Errore', 'Errato', 'Normoudente'].includes(p.tipo);
        if (!isCorrectType) return false;

        let matches = true;

        if (selectedBranch) {
          matches = matches && String(p.store || '').toLowerCase().trim() === selectedBranch.toLowerCase().trim();
        }

        if (selectedPatientType) {
          matches = matches && String(p.tipo || '').toLowerCase().trim() === selectedPatientType.toLowerCase().trim();
        }

        const capToUse = selectedCap || capFilterSearch;
        if (capToUse) {
          matches = matches && String(p.cap || '').trim().includes(capToUse.trim());
        }

        if (selectedStoreId) {
          const store = storesData.find(s => s.id === selectedStoreId);
          if (store) {
            const associatedCaps = storeCaps
              .filter(sc => sc.storeId === selectedStoreId)
              .map(sc => sc.cap);
            const allStoreCaps = [store.cap, ...associatedCaps].filter(Boolean);
            
            if (allStoreCaps.length > 0) {
              matches = matches && allStoreCaps.includes(p.cap);
            }
          }
        }

        if (!selectedBranch && !selectedStoreId && !selectedCap) return false;

        return matches;
      });
      
      const normalizeId = (id: any): string | undefined => {
        if (!id) return undefined;
        const s = String(id).trim();
        if (/^\d+$/.test(s)) {
          return String(parseInt(s, 10));
        }
        return s;
      };

      // Pre-index calls and appointments for O(1) lookup to prevent UI blocking
      const callsById = new Map<string, any[]>();
      const callsByName = new Map<string, any[]>();
      callsData.forEach(c => {
        const id = normalizeId(c.pazienteId);
        if (id) {
          if (!callsById.has(id)) callsById.set(id, []);
          callsById.get(id)!.push(c);
        }
        if (c.nome) {
          const name = c.nome.toLowerCase().trim();
          if (!callsByName.has(name)) callsByName.set(name, []);
          callsByName.get(name)!.push(c);
        }
      });

      const appsById = new Map<string, any[]>();
      const appsByName = new Map<string, any[]>();
      appointmentsData.forEach(a => {
        const id = normalizeId(a.pazienteId);
        if (id) {
          if (!appsById.has(id)) appsById.set(id, []);
          appsById.get(id)!.push(a);
        }
        if (a.nome) {
          const name = a.nome.toLowerCase().trim();
          if (!appsByName.has(name)) appsByName.set(name, []);
          appsByName.get(name)!.push(a);
        }
      });

      const merged = filteredPatients.map(p => {
        const pId = normalizeId(p.id);
        const pName = p.nomeCompleto ? p.nomeCompleto.toLowerCase().trim() : null;

        // Use Set to avoid duplicates if matched by both ID and Name
        const patientCallsSet = new Set<any>();
        if (pId && callsById.has(pId)) callsById.get(pId)!.forEach(c => patientCallsSet.add(c));
        if (pName && callsByName.has(pName)) callsByName.get(pName)!.forEach(c => patientCallsSet.add(c));
        const patientCalls = Array.from(patientCallsSet);

        const patientAppsSet = new Set<any>();
        if (pId && appsById.has(pId)) appsById.get(pId)!.forEach(a => patientAppsSet.add(a));
        if (pName && appsByName.has(pName)) appsByName.get(pName)!.forEach(a => patientAppsSet.add(a));
        const patientApps = Array.from(patientAppsSet);

        // Exclude if future appointment or appointment in the last 2 months
        // OR called in the last 3 months
        const hasFutureApp = patientApps.some(a => (a.esito || '').toLowerCase() === 'appuntamento futuro');
        const hasRecentAppInTwoMonths = patientApps.some(a => a.timestamp >= twoMonthsAgo && a.timestamp <= now.getTime());
        const hasRecentCallInThreeMonths = patientCalls.some(c => c.timestamp >= threeMonthsAgo && c.timestamp <= now.getTime());

        if (hasFutureApp || hasRecentAppInTwoMonths || hasRecentCallInThreeMonths) {
          return null;
        }

        const allNotes = [
          p.note,
          ...patientCalls.map(c => c.note),
          ...patientApps.map(a => a.note)
        ].join(' ').toLowerCase();

        // Exclude if "non interessato" is found in any notes
        if (allNotes.includes('non interessato')) {
          return null;
        }

        let score = 0;
        let reasons: string[] = [];

        const sources = [
          { name: 'Anagrafica', text: p.note },
          ...patientCalls.map(c => ({ name: `Chiamata ${c.data}`, text: c.note })),
          ...patientApps.map(a => ({ name: `Appuntamento ${a.data}`, text: a.note }))
        ];

        const keywordsAids = ['apparecchi', 'protesi', 'sentire meglio', 'prova'];
        const keywordsVisit = ['ipoacusico', 'visita', 'controllo', 'esame', 'udito'];

        sources.forEach(src => {
          if (!src.text) return;
          const lowerText = src.text.toLowerCase();
          
          const foundAids = keywordsAids.filter(k => lowerText.includes(k));
          if (foundAids.length > 0) {
            score += 10;
            reasons.push(`[${src.name}] Interesse apparecchi: "${foundAids.join(', ')}"`);
          }

          const foundVisit = keywordsVisit.filter(k => lowerText.includes(k));
          if (foundVisit.length > 0) {
            score += 5;
            reasons.push(`[${src.name}] Necessità visita: "${foundVisit.join(', ')}"`);
          }
        });

        const sortedCalls = [...patientCalls].sort((a, b) => b.timestamp - a.timestamp);
        const sortedApps = [...patientApps].sort((a, b) => b.timestamp - a.timestamp);

        return {
          ...p,
          score,
          reasons: Array.from(new Set(reasons)), // Deduplicate reasons
          lastCall: sortedCalls.length > 0 ? sortedCalls[0].data : 'Mai',
          lastApp: sortedApps.length > 0 ? sortedApps[0].data : 'Mai'
        };
      }).filter(p => p !== null) as any[];

      const result = merged
        .filter(p => p.score > 0)
        .sort((a, b) => b.score - a.score);

      setAnalyzedPatients(result);
      setIsAnalyzing(false);
    }, 500);
  };

  const handleMaiVistiOChiamati = () => {
    if (patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0) {
      setError("Carica tutti e tre i file (Anagrafica, Telefonate, Appuntamenti) prima di procedere.");
      return;
    }

    if (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim()) {
      setError("Seleziona una filiale, una struttura esterna o un CAP prima di procedere.");
      return;
    }

    setIsAnalyzing(true);
    setError(null);

    setTimeout(() => {
      const filteredPatients = patientsData.filter(p => {
        const isCorrectType = !['Deceduto', 'Cliente', 'Errore', 'Errato', 'Normoudente'].includes(p.tipo);
        if (!isCorrectType) return false;

        let matches = true;

        if (selectedBranch) {
          matches = matches && String(p.store || '').toLowerCase().trim() === selectedBranch.toLowerCase().trim();
        }

        if (selectedPatientType) {
          matches = matches && String(p.tipo || '').toLowerCase().trim() === selectedPatientType.toLowerCase().trim();
        }

        const capToUse = selectedCap || capFilterSearch;
        if (capToUse) {
          matches = matches && String(p.cap || '').trim().includes(capToUse.trim());
        }

        if (selectedStoreId) {
          const store = storesData.find(s => s.id === selectedStoreId);
          if (store) {
            const associatedCaps = storeCaps
              .filter(sc => sc.storeId === selectedStoreId)
              .map(sc => sc.cap);
            const allStoreCaps = [store.cap, ...associatedCaps].filter(Boolean);
            
            if (allStoreCaps.length > 0) {
              matches = matches && allStoreCaps.includes(p.cap);
            }
          }
        }

        if (!selectedBranch && !selectedStoreId && !selectedCap) return false;

        return matches;
      });
      
      const normalizeId = (id: any): string | undefined => {
        if (!id) return undefined;
        const s = String(id).trim();
        if (/^\d+$/.test(s)) {
          return String(parseInt(s, 10));
        }
        return s;
      };

      const callsById = new Map<string, any[]>();
      const callsByName = new Map<string, any[]>();
      callsData.forEach(c => {
        const id = normalizeId(c.pazienteId);
        if (id) {
          if (!callsById.has(id)) callsById.set(id, []);
          callsById.get(id)!.push(c);
        }
        if (c.nome) {
          const name = c.nome.toLowerCase().trim();
          if (!callsByName.has(name)) callsByName.set(name, []);
          callsByName.get(name)!.push(c);
        }
      });

      const appsById = new Map<string, any[]>();
      const appsByName = new Map<string, any[]>();
      appointmentsData.forEach(a => {
        const id = normalizeId(a.pazienteId);
        if (id) {
          if (!appsById.has(id)) appsById.set(id, []);
          appsById.get(id)!.push(a);
        }
        if (a.nome) {
          const name = a.nome.toLowerCase().trim();
          if (!appsByName.has(name)) appsByName.set(name, []);
          appsByName.get(name)!.push(a);
        }
      });

      const result = filteredPatients.filter(p => {
        const pId = normalizeId(p.id);
        const pName = p.nomeCompleto ? p.nomeCompleto.toLowerCase().trim() : null;

        const hasCalls = (pId && callsById.has(pId)) || (pName && callsByName.has(pName));
        const hasApps = (pId && appsById.has(pId)) || (pName && appsByName.has(pName));

        return !hasCalls && !hasApps;
      }).map(p => ({
        ...p,
        score: 0,
        reasons: ["Mai visto o chiamato"],
        lastCall: 'Mai',
        lastApp: 'Mai'
      }));

      setAnalyzedPatients(result);
      setIsAnalyzing(false);
    }, 500);
  };

  const handleResetPatients = () => {
    setSelectedBranch('');
    setSelectedStoreId('');
    setSelectedCap('');
    setSelectedPatientType('');
    setStoreSearch('');
    setCapFilterSearch('');
    setAnalyzedPatients([]);
    setGeocodedPatients([]);
    setGeocodedStores([]);
    setGeocodingProgress({ current: 0, total: 0 });
    setShowMap(false);
    setError(null);
  };

  const handleGeocodePatients = async () => {
    if (patientsData.length === 0) {
      setError("Carica i dati dei pazienti (ii.xlsx) prima di procedere con la mappatura.");
      return;
    }

    if (storesData.length === 0) {
      setError("Avviso: Nessun dato per le strutture (store.xlsx) caricato. La mappa mostrerà solo i pazienti.");
    }

    setShowMap(true);
    setError(null);
    setGeocodedPatients([]);
    setGeocodingProgress({ current: 0, total: 0 });

    // Geocode stores if not already done
    const excludedStores = ["Roma - Sede", "Roma - Villa Chigi", "Roma - San Giovanni"];
    const validStores = storesData.filter(s => {
      const type = String(s.tipo || '').toLowerCase().trim();
      const name = String(s.nome || '').trim();
      return (type.includes('filiale') || type === 'recapito' || type.includes('corporate') || type.includes('franchising')) && !excludedStores.includes(name);
    });

    if (validStores.length > 0 && geocodedStores.length < validStores.length) {
      setIsGeocoding(true);
      const storesToGeocode = validStores.filter(s => !geocodedStores.some(gs => gs.id === s.id));
      let count = geocodedStores.length;
      setGeocodingProgress({ current: count, total: validStores.length });

      for (const s of storesToGeocode) {
        const addressStr = `${s.indirizzo}, ${s.cap} ${s.citta}, Italy`;
        try {
          const response = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(addressStr)}&limit=1`, {
            headers: {
              'User-Agent': 'NoahAnalysisApp/1.0'
            }
          });
          const data = await response.json();
          if (data && data.length > 0) {
            const geocodedStore = { 
              ...s, 
              position: { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) } 
            };
            setGeocodedStores(prev => [...prev, geocodedStore]);
          }
          await new Promise(resolve => setTimeout(resolve, 1000));
        } catch (err) {
          console.error(`Geocoding error for store ${s.nome}:`, err);
        }
        count++;
        setGeocodingProgress({ current: count, total: validStores.length });
      }
      setIsGeocoding(false);
    }
  };

  const handleMapSelection = async (layer: any) => {
    setIsGeocoding(true);
    setError(null);
    
    // Get center of selection to identify city
    let centerLat = 0;
    let centerLng = 0;
    if (layer instanceof L.Circle) {
      centerLat = layer.getLatLng().lat;
      centerLng = layer.getLatLng().lng;
    } else if (layer instanceof L.Rectangle || layer instanceof L.Polygon) {
      centerLat = layer.getBounds().getCenter().lat;
      centerLng = layer.getBounds().getCenter().lng;
    }

    // Try to get city name from reverse geocoding
    let targetCity = '';
    try {
      const revResponse = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${centerLat}&lon=${centerLng}`, {
        headers: { 'User-Agent': 'AskoltaOraAnalysis/1.0' }
      });
      const revData = await revResponse.json();
      targetCity = revData.address?.city || revData.address?.town || revData.address?.village || '';
    } catch (e) {
      console.error("Reverse geocoding failed", e);
    }

    // Filter patients by city and active filters
    let candidates = patientsData.filter(p => {
      let matches = true;
      if (selectedBranch) matches = matches && String(p.store || '').toLowerCase().trim() === selectedBranch.toLowerCase().trim();
      if (selectedPatientType) matches = matches && String(p.tipo || '').toLowerCase().trim() === selectedPatientType.toLowerCase().trim();
      if (selectedCap) matches = matches && String(p.cap || '').trim() === selectedCap.trim();
      if (selectedStoreId) {
        const store = storesData.find(s => s.id === selectedStoreId);
        if (store) {
          const associatedCaps = storeCaps.filter(sc => sc.storeId === selectedStoreId).map(sc => sc.cap);
          const allStoreCaps = [store.cap, ...associatedCaps].filter(Boolean);
          if (allStoreCaps.length > 0) matches = matches && allStoreCaps.includes(p.cap);
        }
      }
      if (targetCity) matches = matches && String(p.citta || '').toLowerCase().includes(targetCity.toLowerCase());
      return matches;
    });

    // If still too many, limit to 200 for selection
    const dataToProcess = candidates.slice(0, 200);
    setGeocodingProgress({ current: 0, total: dataToProcess.length });

    const results: any[] = [];
    const cache = new Map<string, { lat: number, lng: number }>();
    let count = 0;

    for (const p of dataToProcess) {
      const addressStr = `${p.indirizzo}, ${p.cap} ${p.citta}, Italy`;
      let pos = null;

      if (cache.has(addressStr)) {
        pos = cache.get(addressStr);
      } else {
        try {
          const response = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(addressStr)}&limit=1`, {
            headers: { 'User-Agent': 'AskoltaOraAnalysis/1.0' }
          });
          const data = await response.json();
          if (data && data.length > 0) {
            pos = { lat: parseFloat(data[0].lat), lng: parseFloat(data[0].lon) };
            cache.set(addressStr, pos);
          }
          await new Promise(resolve => setTimeout(resolve, 1000));
        } catch (err) {
          console.error(`Geocoding error for ${addressStr}:`, err);
        }
      }

      if (pos && isPointInLayer(pos.lat, pos.lng, layer)) {
        results.push({ ...p, position: pos });
      }
      
      count++;
      setGeocodingProgress({ current: count, total: dataToProcess.length });
    }

    if (results.length > 0) {
      const processed = results.map(p => {
        const pCalls = callsData.filter(c => String(c.pazienteId || '').trim() === String(p.id || '').trim());
        const pApps = appointmentsData.filter(a => String(a.pazienteId || '').trim() === String(p.id || '').trim());
        
        const lastCall = pCalls.length > 0 ? pCalls.sort((a, b) => b.timestamp - a.timestamp)[0].fullDate : 'Mai';
        const lastApp = pApps.length > 0 ? pApps.sort((a, b) => b.timestamp - a.timestamp)[0].fullDate : 'Mai';

        return {
          ...p,
          lastCall,
          lastApp,
          reasons: ['Selezione da Mappa'],
          score: 100
        };
      });

      setAnalyzedPatients(processed);
      setGeocodedPatients(processed);
      // setShowMap(false); // Keep map open
      setError(`Selezionati ${processed.length} pazienti nell'area.`);
    } else {
      setError("Nessun paziente trovato nell'area selezionata.");
    }
    
    // Remove the selection layer from the map to clean up view
    if (layer && typeof layer.remove === 'function') {
      layer.remove();
    }
    
    setIsGeocoding(false);
  };

  const handleNoContattatiDa = () => {
    if (patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0) {
      setError("Carica tutti e tre i file (Anagrafica, Telefonate, Appuntamenti) prima di procedere.");
      return;
    }

    if (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim()) {
      setError("Seleziona una filiale, una struttura esterna o un CAP prima di procedere.");
      return;
    }

    setIsAnalyzing(true);
    setError(null);

    try {
      const result = getNoContattatiPatients(selectedBranch, selectedStoreId, selectedMonths, selectedPatientType, selectedCap || capFilterSearch);
      setAnalyzedPatients(result);
      setIsAnalyzing(false);
    } catch (e) {
      console.error(e);
      setError("Errore durante l'analisi.");
      setIsAnalyzing(false);
    }
  };

  const mapRef = useRef<HTMLDivElement>(null);

  const handleDownloadMap = async () => {
    if (!mapRef.current) return;
    
    setIsGeocoding(true);
    try {
      // Small delay to ensure markers are rendered if just selected
      await new Promise(resolve => setTimeout(resolve, 500));
      
      const dataUrl = await toPng(mapRef.current, {
        cacheBust: true,
        backgroundColor: '#fff',
        filter: (node) => {
          const className = node.className || '';
          if (typeof className === 'string' && className.includes('leaflet-control-container')) return false;
          return true;
        }
      });
      const link = document.createElement('a');
      link.download = `mappa_pazienti_${new Date().getTime()}.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('Error capturing map:', err);
      setError("Impossibile catturare l'immagine della mappa. Verifica la connessione.");
    } finally {
      setIsGeocoding(false);
    }
  };

  const handleSelectRecapitiByAvailability = () => {
    if (availabilityData.length === 0) {
      setError("Carica il file delle disponibilità (dd.xlsx) prima di procedere.");
      return;
    }

    const now = new Date();
    let targetMonth: number;
    let targetYear: number;

    if (availabilityMonth === 'current') {
      targetMonth = now.getMonth();
      targetYear = now.getFullYear();
    } else {
      targetMonth = now.getMonth() + 1;
      targetYear = now.getFullYear();
      if (targetMonth > 11) {
        targetMonth = 0;
        targetYear += 1;
      }
    }

    const eventsInMonth = availabilityData.filter(item => {
      if (!item.timestamp) return false;
      const d = new Date(item.timestamp);
      return d.getMonth() === targetMonth && d.getFullYear() === targetYear;
    });

    const matchedStoresWithDates = storesData
      .filter(s => s.tipo === 'Recapito')
      .map(s => {
        const normalizedStoreName = String(s.nome).toLowerCase().replace(/[^a-z0-9]/g, '').trim();
        
        const storeEvents = eventsInMonth.filter(e => {
          const sn = String(e.storeName).toLowerCase().replace(/[^a-z0-9]/g, '').trim();
          if (!sn) return false;
          
          // Strict matching: exact match after normalization is preferred
          if (sn === normalizedStoreName) return true;
          
          // If not exact, we only allow it if they are very similar in length.
          // This prevents "Civitanova Marche" from matching "Civitanova Marche - Centro Medico La Fenice"
          // but allows minor differences like "Store Name" vs "Store Name -"
          const diff = Math.abs(sn.length - normalizedStoreName.length);
          if (diff <= 5) {
            return sn.includes(normalizedStoreName) || normalizedStoreName.includes(sn);
          }
          
          return false;
        });
        
        // Remove duplicates by date string and sort by timestamp
        const uniqueDatesMap = new Map<string, number>();
        storeEvents.forEach(e => {
          if (!uniqueDatesMap.has(e.date)) {
            uniqueDatesMap.set(e.date, e.timestamp);
          }
        });
        
        const sortedDates = Array.from(uniqueDatesMap.entries())
          .sort((a, b) => a[1] - b[1])
          .map(entry => entry[0]);

        return { id: s.id, dates: sortedDates };
      })
      .filter(s => s.dates.length > 0);

    const matchedStoreIds = matchedStoresWithDates.map(s => s.id);
    const datesMap = matchedStoresWithDates.reduce((acc, curr) => {
      acc[curr.id] = curr.dates.join(', ');
      return acc;
    }, {} as Record<string, string>);

    if (matchedStoreIds.length === 0) {
      setError(`Nessun recapito trovato per il mese ${availabilityMonth === 'current' ? 'corrente' : 'successivo'}.`);
    } else {
      setSelectedRecapitiIds(matchedStoreIds);
      setRecapitiDates(datesMap);
    }
  };

  const enrichPatientsWithData = (patients: any[]) => {
    const normalizeId = (id: any): string | undefined => {
      if (!id) return undefined;
      const s = String(id).trim();
      if (/^\d+$/.test(s)) return String(parseInt(s, 10));
      return s;
    };

    const callsById = new Map<string, any[]>();
    const callsByName = new Map<string, any[]>();
    callsData.forEach(c => {
      const id = normalizeId(c.pazienteId);
      if (id) {
        if (!callsById.has(id)) callsById.set(id, []);
        callsById.get(id)!.push(c);
      }
      if (c.nome) {
        const name = c.nome.toLowerCase().trim();
        if (!callsByName.has(name)) callsByName.set(name, []);
        callsByName.get(name)!.push(c);
      }
    });

    const appsById = new Map<string, any[]>();
    const appsByName = new Map<string, any[]>();
    appointmentsData.forEach(a => {
      const id = normalizeId(a.pazienteId);
      if (id) {
        if (!appsById.has(id)) appsById.set(id, []);
        appsById.get(id)!.push(a);
      }
      if (a.nome) {
        const name = a.nome.toLowerCase().trim();
        if (!appsByName.has(name)) appsByName.set(name, []);
        appsByName.get(name)!.push(a);
      }
    });

    return patients.map(p => {
      const pId = normalizeId(p.id);
      const pName = p.nomeCompleto ? p.nomeCompleto.toLowerCase().trim() : null;
      const calls = [...(pId ? (callsById.get(pId) || []) : []), ...(pName ? (callsByName.get(pName) || []) : [])];
      const apps = [...(pId ? (appsById.get(pId) || []) : []), ...(pName ? (appsByName.get(pName) || []) : [])];
      
      const lastCallTimestamp = calls.reduce((max: number, c) => (c.timestamp && c.timestamp > max) ? c.timestamp : max, 0);
      const lastAppTimestamp = apps.reduce((max: number, a) => (a.timestamp && a.timestamp > max) ? a.timestamp : max, 0);
      
      return {
        ...p,
        lastCall: lastCallTimestamp ? new Date(lastCallTimestamp).toLocaleDateString('it-IT') : 'Mai',
        lastApp: lastAppTimestamp ? new Date(lastAppTimestamp).toLocaleDateString('it-IT') : 'Mai'
      };
    });
  };

  const findDuplicates = () => {
    const EXCLUDED_STORES = ['cuneo', 'grosseto', 'mondovi', 'mondovì', 'alba', 'cassino', "l'aquila", 'avezzano'];
    const groups = new Map<string, any[]>();
    patientsData.forEach(p => {
      const storeName = String(p.store || '').toLowerCase().trim();
      if (EXCLUDED_STORES.some(ex => storeName.includes(ex))) return;

      const nome = String(p.nome || '').toLowerCase().trim();
      const cognome = String(p.cognome || '').toLowerCase().trim();
      const telefono = String(p.telefono || '').trim();
      if (!nome || !cognome || !telefono) return;
      const key = `${nome}|${cognome}|${telefono}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key)!.push(p);
    });
    const results = Array.from(groups.values()).filter(g => g.length > 1).flat();
    setVerificationResults(enrichPatientsWithData(results));
    setActiveVerification('doppioni');
  };

  const findNoCap = () => {
    const EXCLUDED_STORES = ['cuneo', 'grosseto', 'mondovi', 'mondovì', 'alba', 'cassino', "l'aquila", 'avezzano'];
    const results = patientsData.filter(p => {
      const storeName = String(p.store || '').toLowerCase().trim();
      if (EXCLUDED_STORES.some(ex => storeName.includes(ex))) return false;

      const hasAddress = p.indirizzo && p.indirizzo !== 'N/D' && p.indirizzo.trim() !== '';
      const hasCity = p.citta && p.citta !== 'N/D' && p.citta.trim() !== '';
      const noCap = !p.cap || p.cap === 'N/D' || p.cap.trim() === '';
      return hasAddress && hasCity && noCap;
    });
    setVerificationResults(enrichPatientsWithData(results));
    setActiveVerification('nocap');
  };

  const findNoPhone = () => {
    const EXCLUDED_STORES = ['cuneo', 'grosseto', 'mondovi', 'mondovì', 'alba', 'cassino', "l'aquila", 'avezzano'];
    const results = patientsData.filter(p => {
      const storeName = String(p.store || '').toLowerCase().trim();
      if (EXCLUDED_STORES.some(ex => storeName.includes(ex))) return false;
      return !p.telefono || p.telefono === 'N/D' || p.telefono.trim() === '';
    });
    setVerificationResults(enrichPatientsWithData(results));
    setActiveVerification('nophone');
  };

  const findClientsNoProforma = () => {
    const EXCLUDED_STORES = ['cuneo', 'grosseto', 'mondovi', 'mondovì', 'alba', 'cassino', "l'aquila", 'avezzano'];
    const salesCustomerNames = new Set(salesData.map(s => String(s.cliente).toLowerCase().trim()));
    const results = patientsData.filter(p => {
      const storeName = String(p.store || '').toLowerCase().trim();
      if (EXCLUDED_STORES.some(ex => storeName.includes(ex))) return false;

      const tipo = String(p.tipo || '').toLowerCase().trim();
      const isCliente = tipo === 'cliente';
      const hasNoSale = !salesCustomerNames.has(String(p.nomeCompleto || '').toLowerCase().trim());
      return isCliente && hasNoSale;
    });
    setVerificationResults(enrichPatientsWithData(results));
    setActiveVerification('noproforma');
  };

  const handleDownloadRecapitiLists = async () => {
    if (patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0) {
      setError("Carica tutti e tre i file (Anagrafica, Telefonate, Appuntamenti) prima di procedere.");
      return;
    }

    if (selectedRecapitiIds.length === 0) {
      setError("Seleziona almeno un recapito.");
      return;
    }

    setIsAnalyzing(true);
    setError(null);

    try {
      for (const storeId of selectedRecapitiIds) {
        const store = storesData.find(s => s.id === storeId);
        if (!store) continue;
        
        const patients = getNoContattatiPatients('', storeId, selectedMonths, selectedPatientType, '');
        const filename = `${store.nome}-noContattatiDa${selectedMonths}mesi.xlsx`;
        await generateAndDownloadExcel(patients, filename, true);
      }
    } catch (err) {
      console.error(err);
      setError("Errore durante il download delle liste.");
    } finally {
      setIsAnalyzing(false);
    }
  };

  const findContattiChiamati = () => {
    const excludedStores = ['alba', 'cuneo', 'saluzzo', 'mondovi', 'mondovì', 'grosseto', 'savigliano', 'castel del piano', 'orbetello'];
    const results = patientsData.filter(p => {
      const tipo = String(p.tipo || '').toLowerCase().trim();
      const store = String(p.store || '').toLowerCase().trim();
      const isExcluded = excludedStores.some(s => store.includes(s));
      return tipo === 'contatto chiamato' && !isExcluded;
    });

    const normalizeId = (id: any): string | undefined => {
      if (!id) return undefined;
      const s = String(id).trim();
      if (/^\d+$/.test(s)) return String(parseInt(s, 10));
      return s;
    };

    const callsById = new Map<string, any[]>();
    const callsByName = new Map<string, any[]>();
    callsData.forEach(c => {
      const id = normalizeId(c.pazienteId);
      if (id) {
        if (!callsById.has(id)) callsById.set(id, []);
        callsById.get(id)!.push(c);
      }
      if (c.nome) {
        const name = c.nome.toLowerCase().trim();
        if (!callsByName.has(name)) callsByName.set(name, []);
        callsByName.get(name)!.push(c);
      }
    });

    const appsById = new Map<string, any[]>();
    const appsByName = new Map<string, any[]>();
    appointmentsData.forEach(a => {
      const id = normalizeId(a.pazienteId);
      if (id) {
        if (!appsById.has(id)) appsById.set(id, []);
        appsById.get(id)!.push(a);
      }
      if (a.nome) {
        const name = a.nome.toLowerCase().trim();
        if (!appsByName.has(name)) appsByName.set(name, []);
        appsByName.get(name)!.push(a);
      }
    });

    const enrichedResults = enrichPatientsWithData(results);

    const processedResults = enrichedResults.map(p => {
      const pId = normalizeId(p.id);
      const pName = p.nomeCompleto ? p.nomeCompleto.toLowerCase().trim() : null;
      
      const patientCalls = [...(pId ? (callsById.get(pId) || []) : []), ...(pName ? (callsByName.get(pName) || []) : [])];
      const patientApps = [...(pId ? (appsById.get(pId) || []) : []), ...(pName ? (appsByName.get(pName) || []) : [])];
      
      const allNotes = [
        p.note,
        ...patientCalls.map(c => c.note),
        ...patientApps.map(a => a.note)
      ].filter(Boolean).join(' ').toLowerCase();

      let classification = '';
      let classificationReason = '';
      
      const normoMatch = allNotes.match(/normoudente|nessuna perdita|udito normale|ci sente bene|audito ok/i);
      const pcMatch = allNotes.match(/\bpc\b|possibile cliente|interessat/i);
      const premMatch = allNotes.match(/prematuro|lieve perdita|poca perdita|inizio perdita/i);

      if (normoMatch) {
        classification = 'Normoudente';
        classificationReason = `Trovato riferimento: "${normoMatch[0]}"`;
      } else if (pcMatch) {
        classification = 'PC';
        classificationReason = `Trovato riferimento: "${pcMatch[0]}"`;
      } else if (premMatch) {
        classification = 'Prematuro';
        classificationReason = `Trovato riferimento: "${premMatch[0]}"`;
      }

      return {
        ...p,
        classification,
        classificationReason
      };
    });

    processedResults.sort((a, b) => {
      const storeA = String(a.store || '').toLowerCase().trim();
      const storeB = String(b.store || '').toLowerCase().trim();
      return storeA.localeCompare(storeB);
    });

    setVerificationResults(processedResults);
    setActiveVerification('contattichiamati');
  };

  const handleMarketingAnalysis = () => {
    if (salesData.length === 0) {
      setError("Carica il file delle vendite (ff.xlsx) prima di procedere.");
      return;
    }

    const filteredSales = salesData.filter(s => s.year === selectedYear);
    const filteredCredits = creditNotes.filter(c => c.year === selectedYear);
    
    // Group by month, then by channel
    const monthlyData: any[] = [];
    
    for (let m = 0; m < 12; m++) {
      const monthSales = filteredSales.filter(s => s.month === m);
      const monthCredits = filteredCredits.filter(c => c.month === m);
      const channels = new Map<string, { total: number, count: number }>();
      
      monthSales.forEach(s => {
        const channel = s.canale || 'N/D';
        if (!channels.has(channel)) {
          channels.set(channel, { total: 0, count: 0 });
        }
        const stats = channels.get(channel)!;
        stats.total += s.valore;
        stats.count += 1;
      });

      monthCredits.forEach(c => {
        const channel = c.canale || 'N/D';
        if (!channels.has(channel)) {
          channels.set(channel, { total: 0, count: 0 });
        }
        const stats = channels.get(channel)!;
        stats.total -= c.valore;
        // We don't subtract from count, as count is for sales
      });
      
      if (channels.size > 0) {
        monthlyData.push({
          month: MONTHS[m],
          channels: Array.from(channels.entries())
            .map(([name, stats]) => ({ name, ...stats }))
            .sort((a, b) => b.total - a.total)
        });
      }
    }
    
    setMarketingAnalysis(monthlyData);
    setShowMarketingAnalysis(true);
  };

  const handleScreeningAnalysis = () => {
    if (salesData.length === 0) {
      setError("Carica il file delle vendite (ff.xlsx) prima di procedere.");
      return;
    }

    const filteredSales = salesData.filter(s => 
      s.year === selectedYear && 
      (s.canale?.toUpperCase() === 'SCREENING' || s.canale?.toUpperCase() === 'PROMOTER')
    );
    
    const filteredCredits = creditNotes.filter(c => 
      c.year === selectedYear && 
      (c.canale?.toUpperCase() === 'SCREENING' || c.canale?.toUpperCase() === 'PROMOTER')
    );

    // Group by Deposito di origine, then by Canale Secondario
    const stats: Record<string, { channels: Record<string, number[]> }> = {};

    filteredSales.forEach(sale => {
      const deposito = sale.depositoOrigine || 'N/D';
      const secondaryChannel = sale.canaleSecondario || 'N/D';
      
      if (!stats[deposito]) {
        stats[deposito] = { channels: {} };
      }
      if (!stats[deposito].channels[secondaryChannel]) {
        stats[deposito].channels[secondaryChannel] = new Array(12).fill(0);
      }
      stats[deposito].channels[secondaryChannel][sale.month] += sale.valore;
    });

    filteredCredits.forEach(note => {
      const deposito = note.depositoOrigine || 'N/D';
      const secondaryChannel = note.canaleSecondario || 'N/D';
      
      if (!stats[deposito]) {
        stats[deposito] = { channels: {} };
      }
      if (!stats[deposito].channels[secondaryChannel]) {
        stats[deposito].channels[secondaryChannel] = new Array(12).fill(0);
      }
      stats[deposito].channels[secondaryChannel][note.month] -= note.valore;
    });

    const processedData = Object.entries(stats).map(([depositoName, depData]) => ({
      name: depositoName,
      channels: Object.entries(depData.channels).map(([cName, cMonthly]) => ({
        name: cName,
        monthly: cMonthly,
        total: cMonthly.reduce((a, b) => a + b, 0)
      })).sort((a, b) => b.total - a.total),
      total: Object.values(depData.channels).reduce((acc, monthly) => acc + monthly.reduce((a, b) => a + b, 0), 0)
    })).sort((a, b) => a.name.localeCompare(b.name));

    setScreeningAnalysis(processedData);
    setShowScreeningAnalysis(true);
  };

  const handleMediciAnalysis = () => {
    if (salesData.length === 0) {
      setError("Carica il file delle vendite (ff.xlsx) prima di procedere.");
      return;
    }

    const filteredSales = salesData.filter(s => 
      s.year === selectedYear && 
      s.canale?.toUpperCase() === 'MEDICO'
    );
    
    const filteredCredits = creditNotes.filter(c => 
      c.year === selectedYear && 
      c.canale?.toUpperCase() === 'MEDICO'
    );

    // Group by Audioprotesista, then by Canale Secondario
    const stats: Record<string, { channels: Record<string, number[]> }> = {};

    filteredSales.forEach(sale => {
      const audioprotesista = sale.audioprotesista || 'N/D';
      const secondaryChannel = sale.canaleSecondario || 'N/D';
      
      if (!stats[audioprotesista]) {
        stats[audioprotesista] = { channels: {} };
      }
      if (!stats[audioprotesista].channels[secondaryChannel]) {
        stats[audioprotesista].channels[secondaryChannel] = new Array(12).fill(0);
      }
      stats[audioprotesista].channels[secondaryChannel][sale.month] += sale.valore;
    });

    filteredCredits.forEach(note => {
      const audioprotesista = note.audioprotesista || 'N/D';
      const secondaryChannel = note.canaleSecondario || 'N/D';
      
      if (!stats[audioprotesista]) {
        stats[audioprotesista] = { channels: {} };
      }
      if (!stats[audioprotesista].channels[secondaryChannel]) {
        stats[audioprotesista].channels[secondaryChannel] = new Array(12).fill(0);
      }
      stats[audioprotesista].channels[secondaryChannel][note.month] -= note.valore;
    });

    const processedData = Object.entries(stats).map(([audioprotesistaName, depData]) => ({
      name: audioprotesistaName,
      channels: Object.entries(depData.channels).map(([cName, cMonthly]) => ({
        name: cName,
        monthly: cMonthly,
        total: cMonthly.reduce((a, b) => a + b, 0)
      })).sort((a, b) => b.total - a.total),
      total: Object.values(depData.channels).reduce((acc, monthly) => acc + monthly.reduce((a, b) => a + b, 0), 0)
    })).sort((a, b) => a.name.localeCompare(b.name));

    setMediciAnalysis(processedData);
    setShowMediciAnalysis(true);
  };

  const handleDownloadScreeningExcel = async () => {
    if (screeningAnalysis.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Analisi Screening ${selectedYear}`);

    // Define columns
    const columns = [
      { header: 'Deposito di origine / Canale Secondario', key: 'name', width: 45 },
      { header: 'Totale Anno', key: 'total', width: 15 },
      ...MONTHS.map((m, idx) => ({ header: m, key: `month_${idx}`, width: 12 }))
    ];

    worksheet.columns = columns;
    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    // Style header
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });

    let currentRowIndex = 2;

    screeningAnalysis.forEach(dep => {
      // Add Deposito Row
      const depRowData: any = {
        name: dep.name,
        total: dep.total,
      };
      MONTHS.forEach((_, mIdx) => {
        depRowData[`month_${mIdx}`] = dep.channels.reduce((acc: number, c: any) => acc + c.monthly[mIdx], 0);
      });

      const depRow = worksheet.addRow(depRowData);
      depRow.font = { bold: true };
      const depBgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
      depRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: depBgColor } };
        cell.numFmt = '#,##0';
      });
      currentRowIndex++;

      // Add Channel Rows
      dep.channels.forEach((c: any) => {
        const cRowData: any = {
          name: `  ↳ ${c.name}`,
          total: c.total,
        };
        MONTHS.forEach((_, mIdx) => {
          cRowData[`month_${mIdx}`] = c.monthly[mIdx];
        });

        const cRow = worksheet.addRow(cRowData);
        const cBgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
        cRow.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cBgColor } };
          cell.numFmt = '#,##0';
        });
        currentRowIndex++;
      });
    });

    // Add Grand Total
    const totalsData: any = {
      name: 'TOTALE GENERALE SCREENING',
      total: screeningAnalysis.reduce((acc, ap) => acc + ap.total, 0),
    };
    MONTHS.forEach((_, mIdx) => {
      totalsData[`month_${mIdx}`] = screeningAnalysis.reduce((acc, ap) => 
        acc + ap.channels.reduce((cAcc: number, c: any) => cAcc + c.monthly[mIdx], 0), 0);
    });

    const totalRow = worksheet.addRow(totalsData);
    totalRow.font = { bold: true };
    totalRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
      cell.numFmt = '#,##0';
      cell.border = { top: { style: 'medium' }, bottom: { style: 'medium' } };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Analisi_Screening_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const handleDownloadMediciExcel = async () => {
    if (mediciAnalysis.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Analisi Medici ${selectedYear}`);

    // Define columns
    const columns = [
      { header: 'Audioprotesista / Canale Secondario', key: 'name', width: 45 },
      { header: 'Totale Anno', key: 'total', width: 15 },
      ...MONTHS.map((m, idx) => ({ header: m, key: `month_${idx}`, width: 12 }))
    ];

    worksheet.columns = columns;
    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    // Style header
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });

    let currentRowIndex = 2;

    mediciAnalysis.forEach(dep => {
      // Add Audioprotesista Row
      const depRowData: any = {
        name: dep.name,
        total: dep.total,
      };
      MONTHS.forEach((_, mIdx) => {
        depRowData[`month_${mIdx}`] = dep.channels.reduce((acc: number, c: any) => acc + c.monthly[mIdx], 0);
      });

      const depRow = worksheet.addRow(depRowData);
      depRow.font = { bold: true };
      const depBgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
      depRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: depBgColor } };
        cell.numFmt = '#,##0';
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      });
      currentRowIndex++;

      // Add Channel Rows
      dep.channels.forEach((c: any) => {
        const channelRowData: any = {
          name: `   ↳ ${c.name}`,
          total: c.total,
        };
        MONTHS.forEach((_, mIdx) => {
          channelRowData[`month_${mIdx}`] = c.monthly[mIdx];
        });

        const channelRow = worksheet.addRow(channelRowData);
        channelRow.font = { italic: true };
        const channelBgColor = currentRowIndex % 2 === 0 ? 'FFF2F2F2' : 'FFFFFFFF';
        channelRow.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: channelBgColor } };
          cell.numFmt = '#,##0';
          cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        });
        currentRowIndex++;
      });
    });

    // Add Totals Row
    const totalsData: any = {
      name: 'TOTALE GENERALE MEDICI',
      total: mediciAnalysis.reduce((acc, ap) => acc + ap.total, 0),
    };
    MONTHS.forEach((_, mIdx) => {
      totalsData[`month_${mIdx}`] = mediciAnalysis.reduce((acc, ap) => 
        acc + ap.channels.reduce((cAcc: number, c: any) => cAcc + c.monthly[mIdx], 0), 0);
    });

    const totalRow = worksheet.addRow(totalsData);
    totalRow.font = { bold: true };
    totalRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
      cell.numFmt = '#,##0';
      cell.border = { top: { style: 'medium' }, bottom: { style: 'medium' } };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Analisi_Medici_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const handleCoefficientiAnalysis = () => {
    if (salesData.length === 0) {
      setError("Carica il file delle vendite (ff.xlsx) prima di procedere.");
      return;
    }

    const filteredYearSales = salesData.filter(s => s.year === selectedYear);
    const filteredYearCredits = creditNotes.filter(c => c.year === selectedYear);
    
    // Group by Audioprotesista
    const statsByAP: Record<string, { totalValore: number, totalQuantita: number, customers: Set<string> }> = {};

    filteredYearSales.forEach(s => {
      const ap = s.audioprotesista || 'N/D';
      if (!statsByAP[ap]) {
        statsByAP[ap] = { totalValore: 0, totalQuantita: 0, customers: new Set() };
      }
      statsByAP[ap].totalValore += (s.valore || 0);
      statsByAP[ap].totalQuantita += (s.quantita || 0);
      if (s.cliente) {
        statsByAP[ap].customers.add(s.cliente.toLowerCase().trim());
      }
    });

    const processed = Object.entries(statsByAP).map(([name, data]) => {
      const apCredits = filteredYearCredits.filter(c => (c.audioprotesista || 'N/D') === name);
      const totalCreditsVal = apCredits.reduce((sum, c) => sum + (c.valore || 0), 0);
      const netValore = data.totalValore - totalCreditsVal;

      const numCustomers = data.customers.size || 1;
      const mediaVendita = data.totalQuantita > 0 ? netValore / data.totalQuantita : 0;
      const coefficienteBinaurale = data.totalQuantita / numCustomers;
      
      return {
        name,
        netValore,
        totalValore: data.totalValore,
        totalCreditsVal,
        totalQuantita: data.totalQuantita,
        numCustomers,
        mediaVendita,
        coefficienteBinaurale
      };
    }).sort((a, b) => b.netValore - a.netValore);

    setCoefficientiAnalysis(processed);
    setShowCoefficientiAnalysis(true);
  };

  const handleDownloadSalesExcel = async () => {
    if (aggregatedData.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Vendite ${selectedYear}`);

    // Define columns
    const columns = [
      { header: 'Audioprotesista', key: 'name', width: 35 },
      { header: 'Totale Anno', key: 'total', width: 15 },
      ...MONTHS.map((m, idx) => ({ header: m, key: `month_${idx}`, width: 12 }))
    ];

    worksheet.columns = columns;

    // Freeze the first row
    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    // Style the header row
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' } // Blue background
      };
      cell.font = {
        color: { argb: 'FFFFFFFF' }, // White text
        bold: true
      };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF376091' } },
        left: { style: 'thin', color: { argb: 'FF376091' } },
        bottom: { style: 'thin', color: { argb: 'FF376091' } },
        right: { style: 'thin', color: { argb: 'FF376091' } }
      };
    });

    let currentRowIndex = 2;

    (aggregatedData as any[]).forEach(([name, data]) => {
      const rowNetTotal = (data.sales as number[]).reduce((a, b) => a + b, 0) - (data.credits as number[]).reduce((a, b) => a + b, 0);
      
      const mainRowData: any = {
        name: name,
        total: rowNetTotal,
      };
      MONTHS.forEach((_, idx) => {
        mainRowData[`month_${idx}`] = (data.sales[idx] as number) - (data.credits[idx] as number);
      });

      const mainRow = worksheet.addRow(mainRowData);
      mainRow.font = { bold: true };
      
      // Apply alternating background color
      const bgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
      mainRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgColor }
        };
        cell.numFmt = '#,##0';
        cell.border = {
          left: { style: 'thin', color: { argb: 'FFB8CCE4' } },
          right: { style: 'thin', color: { argb: 'FFB8CCE4' } }
        };
      });
      currentRowIndex++;

      if (showAllDetails) {
        Object.entries(data.channelMonthly as Record<string, number[]>).forEach(([channel, monthlyValues]) => {
          const channelTotal = (monthlyValues as number[]).reduce((a, b) => a + b, 0);
          const channelRowData: any = {
            name: `  ↳ ${channel}`,
            total: channelTotal,
          };
          MONTHS.forEach((_, idx) => {
            channelRowData[`month_${idx}`] = (monthlyValues[idx] as number);
          });

          const channelRow = worksheet.addRow(channelRowData);
          const channelBgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
          channelRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: channelBgColor }
            };
            cell.numFmt = '#,##0';
            cell.border = {
              left: { style: 'thin', color: { argb: 'FFB8CCE4' } },
              right: { style: 'thin', color: { argb: 'FFB8CCE4' } }
            };
          });
          currentRowIndex++;
        });
      }
    });

    // Add total row
    const totalsData: any = {
      name: 'TOTALE GENERALE',
      total: (aggregatedData as any[]).reduce((acc, [_, data]) => 
        acc + (data.sales as number[]).reduce((a, b) => a + b, 0) - (data.credits as number[]).reduce((a, b) => a + b, 0), 0
      ),
    };
    MONTHS.forEach((_, idx) => {
      totalsData[`month_${idx}`] = (aggregatedData as any[]).reduce((acc, [_, data]) => acc + (data.sales[idx] as number) - (data.credits[idx] as number), 0);
    });

    const totalRow = worksheet.addRow(totalsData);
    totalRow.font = { bold: true };
    totalRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDCE6F1' } // Light blue for total
      };
      cell.numFmt = '#,##0';
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF376091' } },
        bottom: { style: 'medium', color: { argb: 'FF376091' } }
      };
    });

    // Generate buffer and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Riepilogo_Vendite_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const downloadAudioPerformanceExcel = async () => {
    if (!audioPerformance || !selectedAudioPro) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Analisi_${selectedAudioPro.slice(0, 20)}`);

    const addTable = (title: string, headers: string[], rows: any[][]) => {
      worksheet.addRow([]);
      const titleRow = worksheet.addRow([title]);
      titleRow.getCell(1).font = { bold: true, size: 14 };
      
      const headerRow = worksheet.addRow(headers);
      headerRow.font = { bold: true, size: 10 };
      headerRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
        cell.border = { 
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        cell.alignment = { horizontal: 'center' };
      });
      headerRow.getCell(1).alignment = { horizontal: 'left' };

      rows.forEach((rowData, idx) => {
        const row = worksheet.addRow(rowData);
        const isTotalRow = idx === rows.length - 1 && (String(rowData[0]).includes('TOTALE') || String(rowData[0]).includes('Saldo'));
        
        row.eachCell((cell, colNumber) => {
          if (colNumber > 1) {
            cell.numFmt = '#,##0';
            cell.alignment = { horizontal: 'center' };
          }
          if (isTotalRow) {
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9FAFB' } };
          }
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFEEEEEE' } },
            left: { style: 'thin', color: { argb: 'FFEEEEEE' } },
            bottom: { style: 'thin', color: { argb: 'FFEEEEEE' } },
            right: { style: 'thin', color: { argb: 'FFEEEEEE' } }
          };
        });
      });
    };

    const monthHeaders = ['Descrizione', ...MONTHS, 'TOTALE'];

    // 1. Events
    const eventRows = audioPerformance.events.map(([store, months]) => [
      store,
      ...months,
      months.reduce((a, b) => a + b, 0)
    ]);
    const eventColTotals = new Array(12).fill(0);
    audioPerformance.events.forEach(([_, m]) => m.forEach((v, i) => eventColTotals[i] += v));
    eventRows.push(['TOTALE EVENTI', ...eventColTotals, eventColTotals.reduce((a, b) => a + b, 0)]);
    addTable('01. TABELLA EVENTI (dd.xlsx)', monthHeaders, eventRows);

    // 2. Calls
    const callRows = audioPerformance.calls.map(([outcome, months]) => [
      outcome,
      ...months,
      months.reduce((a, b) => a + b, 0)
    ]);
    const callColTotals = new Array(12).fill(0);
    audioPerformance.calls.forEach(([_, m]) => m.forEach((v, i) => callColTotals[i] += v));
    callRows.push(['TOTALE TELEFONATE', ...callColTotals, callColTotals.reduce((a, b) => a + b, 0)]);
    addTable('02. TABELLA TELEFONATE (tt.xlsx)', monthHeaders, callRows);

    // 3. Appointments
    const apptRows = audioPerformance.appointments.map(([type, months]) => [
      type,
      ...months,
      months.reduce((a, b) => a + b, 0)
    ]);
    const apptColTotals = new Array(12).fill(0);
    audioPerformance.appointments.forEach(([_, m]) => m.forEach((v, i) => apptColTotals[i] += v));
    apptRows.push(['TOTALE APPUNTAMENTI', ...apptColTotals, apptColTotals.reduce((a, b) => a + b, 0)]);
    addTable('03. TABELLA APPUNTAMENTI (aa.xlsx)', monthHeaders, apptRows);

    // 4. Sales & Credits
    const financeRows = [
      ['Vendite (ff.xlsx)', ...audioPerformance.finance.sales, audioPerformance.finance.sales.reduce((a, b) => a + b, 0)],
      ['Storni (rr.xlsx)', ...audioPerformance.finance.credits, audioPerformance.finance.credits.reduce((a, b) => a + b, 0)],
      ['Saldo Netto', ...audioPerformance.finance.balance, audioPerformance.finance.balance.reduce((a, b) => a + b, 0)]
    ];
    addTable('04. VENDITE E STORNI (ff.xlsx / rr.xlsx)', monthHeaders, financeRows);

    // 5. Open Trials
    if (audioPerformance.openTrials.length > 0) {
      worksheet.addRow([]);
      const tTitle = worksheet.addRow(['05. DETTAGLIO APPARECCHI IN PROVA']);
      tTitle.getCell(1).font = { bold: true, size: 14 };
      const tHeader = worksheet.addRow(['Cliente', 'Valore']);
      tHeader.font = { bold: true };
      tHeader.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } });
      
      audioPerformance.openTrials.forEach(t => {
        const row = worksheet.addRow([t.cliente, t.valore]);
        row.getCell(2).numFmt = '#,##0.00';
      });
      
      const tFooter = worksheet.addRow(['TOTALE PROVE APERTE', audioPerformance.openTrials.reduce((sum, t) => sum + (t.valore || 0), 0)]);
      tFooter.font = { bold: true };
      tFooter.getCell(2).numFmt = '#,##0.00';
    }

    worksheet.getColumn(1).width = 40;
    for (let i = 2; i <= 14; i++) {
        worksheet.getColumn(i).width = 12;
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Analisi_${selectedAudioPro.replace(/\s+/g, '_')}_${selectedAudioYear}.xlsx`;
    link.click();
    window.URL.revokeObjectURL(url);
  };

  const handleDownloadMarketingExcel = async () => {
    if (marketingAnalysis.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Analisi Canali ${selectedYear}`);

    // Define columns
    const columns = [
      { header: 'Canale Primario', key: 'name', width: 35 },
      { header: 'Totale Anno', key: 'total', width: 15 },
      ...MONTHS.map((m, idx) => ({ header: m, key: `month_${idx}`, width: 12 }))
    ];

    worksheet.columns = columns;

    // Freeze the first row
    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    // Style the header row
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' } // Blue background
      };
      cell.font = {
        color: { argb: 'FFFFFFFF' }, // White text
        bold: true
      };
      cell.alignment = { vertical: 'middle', horizontal: 'left' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF376091' } },
        left: { style: 'thin', color: { argb: 'FF376091' } },
        bottom: { style: 'thin', color: { argb: 'FF376091' } },
        right: { style: 'thin', color: { argb: 'FF376091' } }
      };
    });

    let currentRowIndex = 2;

    const channelNames = (Array.from(new Set(marketingAnalysis.flatMap(m => m.channels.map((c: any) => c.name)))) as string[]).sort();

    channelNames.forEach(channelName => {
      const channelTotal = marketingAnalysis.reduce((acc, m) => {
        const c = m.channels.find((ch: any) => ch.name === channelName);
        return acc + (c ? c.total : 0);
      }, 0);

      const rowData: any = {
        name: channelName,
        total: channelTotal,
      };

      MONTHS.forEach((_, mIdx) => {
        const monthData = marketingAnalysis.find(m => m.month === MONTHS[mIdx]);
        const channelData = monthData?.channels.find((c: any) => c.name === channelName);
        rowData[`month_${mIdx}`] = channelData ? channelData.total : 0;
      });

      const row = worksheet.addRow(rowData);
      row.font = { bold: true };
      
      // Apply alternating background color
      const bgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
      row.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: bgColor }
        };
        cell.numFmt = '#,##0';
        cell.border = {
          left: { style: 'thin', color: { argb: 'FFB8CCE4' } },
          right: { style: 'thin', color: { argb: 'FFB8CCE4' } }
        };
      });
      currentRowIndex++;
    });

    // Add total row
    const grandTotal = marketingAnalysis.reduce((acc, m) => acc + m.channels.reduce((cAcc: number, c: any) => cAcc + c.total, 0), 0);
    const totalsData: any = {
      name: 'TOTALE GENERALE',
      total: grandTotal,
    };
    MONTHS.forEach((_, mIdx) => {
      totalsData[`month_${mIdx}`] = marketingAnalysis.find(m => m.month === MONTHS[mIdx])?.channels.reduce((acc: number, c: any) => acc + c.total, 0) || 0;
    });

    const totalRow = worksheet.addRow(totalsData);
    totalRow.font = { bold: true };
    totalRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFDCE6F1' } // Light blue for total
      };
      cell.numFmt = '#,##0';
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF376091' } },
        bottom: { style: 'medium', color: { argb: 'FF376091' } }
      };
    });

    // Generate buffer and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Analisi_Canali_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const handleDownloadDrillDownExcel = async () => {
    if (!drillDown) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dettaglio');

    // Define columns
    const columns = [
      { header: 'Tipo', key: 'tipo', width: 10 },
      { header: 'Audioprotesista', key: 'audioprotesista', width: 25 },
      { header: 'Data', key: 'date', width: 15 },
      { header: 'Cliente finale', key: 'cliente', width: 30 },
      { header: 'Canale', key: 'canale', width: 20 },
      { header: 'Intermediario', key: 'intermediario', width: 25 },
      { header: 'Importo', key: 'importo', width: 15 }
    ];

    worksheet.columns = columns;
    worksheet.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

    // Style header
    const headerRow = worksheet.getRow(1);
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
      cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    });

    let currentRowIndex = 2;

    // Add Sales
    drillDown.sales.forEach((s: any) => {
      const row = worksheet.addRow({
        tipo: 'VENDITA',
        audioprotesista: s.audioprotesista || 'N/D',
        date: s.fullDate,
        cliente: s.cliente,
        canale: s.canale || 'N/D',
        intermediario: s.intermediario || 'N/D',
        importo: s.valore
      });
      const bgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.numFmt = '#,##0';
      });
      currentRowIndex++;
    });

    // Add Credits
    drillDown.credits.forEach((c: any) => {
      const row = worksheet.addRow({
        tipo: 'STORNO',
        audioprotesista: c.audioprotesista || 'N/D',
        date: c.fullDate,
        cliente: c.cliente,
        canale: c.canale || 'N/D',
        intermediario: c.intermediario || 'N/D',
        importo: -c.valore
      });
      const bgColor = currentRowIndex % 2 === 0 ? 'FFDCE6F1' : 'FFFFFFFF';
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.numFmt = '#,##0';
        cell.font = { color: { argb: 'FFFF0000' } }; // Red for credits
      });
      currentRowIndex++;
    });

    // Add Total
    const totalNet = drillDown.sales.reduce((acc: number, s: any) => acc + s.valore, 0) - 
                     drillDown.credits.reduce((acc: number, c: any) => acc + c.valore, 0);
    
    const totalRow = worksheet.addRow({
      tipo: 'TOTALE NETTO',
      importo: totalNet
    });
    totalRow.font = { bold: true };
    totalRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
      cell.numFmt = '#,##0';
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Dettaglio_${drillDown.name.replace(/\s+/g, '_')}_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const downloadAppointmentsExcel = async () => {
    if (appointmentsAnalysis.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Appuntamenti ${selectedApptYear}`);

    const columns = [
      { header: appointmentsAnalysisView === 'canali' ? 'Audioprotesista / Canale' : 'Audioprotesista / Tipo', key: 'label', width: 40 },
      ...MONTH_NAMES.map((m, i) => ({ header: m, key: `m${i}`, width: 12 })),
      { header: 'TOTALE', key: 'total', width: 15 }
    ];

    worksheet.columns = columns;

    // Styling headers
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };
      cell.alignment = { horizontal: 'center' };
    });

    appointmentsAnalysis.forEach(pro => {
      // Add Audioprotesista row
      const proRowData: any = {
        label: pro.name,
        total: pro.months.reduce((a, b) => a + b, 0)
      };
      pro.months.forEach((count, i) => {
        proRowData[`m${i}`] = count || 0;
      });

      const proRow = worksheet.addRow(proRowData);
      proRow.font = { bold: true };
      proRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      });

      // Add Channel rows
      Array.from(pro.subGroups.entries()).forEach(([subGroup, subGroupMonths]) => {
        const chanRowData: any = {
          label: `  ${subGroup}`,
          total: subGroupMonths.reduce((a, b) => a + b, 0)
        };
        subGroupMonths.forEach((count, i) => {
          chanRowData[`m${i}`] = count || 0;
        });

        const chanRow = worksheet.addRow(chanRowData);
        chanRow.font = { italic: true };
      });
    });

    // Add Grand Total row
    const grandTotalRowData: any = {
      label: 'TOTALE GENERALE',
      total: appointmentsAnalysis.reduce((total, pro) => total + pro.months.reduce((a, b) => a + b, 0), 0)
    };
    Array.from({ length: 12 }).forEach((_, mIdx) => {
      grandTotalRowData[`m${mIdx}`] = appointmentsAnalysis.reduce((sum, pro) => sum + (pro.months[mIdx] || 0), 0);
    });

    const grandTotalRow = worksheet.addRow(grandTotalRowData);
    grandTotalRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    grandTotalRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };
      cell.alignment = { horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Analisi_Appuntamenti_${selectedApptYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const downloadScreeningEntrancesExcel = async () => {
    if (screeningEntrancesAnalysis.length === 0) return;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Ingressi Screening ${selectedScreeningYear}`);

    const columns = [
      { header: 'Audioprotesista / Store', key: 'label', width: 40 },
      ...MONTH_NAMES.map((m, i) => ({ header: m, key: `m${i}`, width: 12 })),
      { header: 'TOTALE', key: 'total', width: 15 }
    ];

    worksheet.columns = columns;

    // Styling headers
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };
      cell.alignment = { horizontal: 'center' };
    });

    screeningEntrancesAnalysis.forEach(pro => {
      // Add Audioprotesista row
      const proRowData: any = {
        label: pro.name,
        total: pro.months.reduce((a, b) => a + b, 0)
      };
      pro.months.forEach((count, i) => {
        proRowData[`m${i}`] = count || 0;
      });

      const proRow = worksheet.addRow(proRowData);
      proRow.font = { bold: true };
      proRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      });

      // Add Store rows
      Array.from(pro.stores.entries()).forEach(([store, storeMonths]) => {
        const storeRowData: any = {
          label: `  ${store}`,
          total: storeMonths.reduce((a, b) => a + b, 0)
        };
        storeMonths.forEach((count, i) => {
          storeRowData[`m${i}`] = count || 0;
        });

        const storeRow = worksheet.addRow(storeRowData);
        storeRow.font = { italic: true };
      });
    });

    // Add Grand Total row
    const grandTotalRowData: any = {
      label: 'TOTALE GENERALE',
      total: screeningEntrancesAnalysis.reduce((total, pro) => total + pro.months.reduce((a, b) => a + b, 0), 0)
    };
    Array.from({ length: 12 }).forEach((_, mIdx) => {
      grandTotalRowData[`m${mIdx}`] = screeningEntrancesAnalysis.reduce((sum, pro) => sum + (pro.months[mIdx] || 0), 0);
    });

    const grandTotalRow = worksheet.addRow(grandTotalRowData);
    grandTotalRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    grandTotalRow.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF141414' } };
      cell.alignment = { horizontal: 'center' };
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Analisi_Ingressi_Screening_${selectedScreeningYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const handleDownloadBudgetSituation = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Budget ${selectedYear}`);

    const BORDER_STYLE: Partial<ExcelJS.Borders> = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      left: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'thin', color: { argb: 'FF000000' } },
    };

    const getColName = (n: number) => {
      let abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
      let name = "";
      while (n > 0) {
        let mod = (n - 1) % 26;
        name = abc[mod] + name;
        n = Math.floor((n - mod) / 26);
      }
      return name;
    };

    let currentRow = 1;

    // We'll create blocks for each quarter
    for (let q = 0; q < 4; q++) {
      const monthsInQ = [q * 3, q * 3 + 1, q * 3 + 2];
      const qLabel = ["I TRIMESTRE", "II TRIMESTRE", "III TRIMESTRE", "IV TRIMESTRE"][q];

      // --- Header Row 1 (Yellow) ---
      const header1 = [''];
      monthsInQ.forEach(mIdx => {
        header1.push(MONTHS[mIdx].toLowerCase(), '', '');
      });
      header1.push(qLabel, '', '', ''); // Quarterly header: fatt, bgt, diff, %
      
      const row1 = worksheet.addRow(header1);
      row1.height = 25;
      row1.eachCell((cell, colNumber) => {
        if (colNumber > 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.font = { bold: true };
          cell.border = BORDER_STYLE;
        }
      });

      // Merge Yellow Headers
      for (let i = 0; i < 3; i++) {
        const colStart = 2 + (i * 3);
        worksheet.mergeCells(`${getColName(colStart)}${currentRow}:${getColName(colStart + 2)}${currentRow}`);
      }
      const qColStart = 11;
      worksheet.mergeCells(`${getColName(qColStart)}${currentRow}:${getColName(qColStart + 2)}${currentRow}`);

      currentRow++;

      // --- Header Row 2 (Light Blue) ---
      const header2 = [''];
      for (let i = 0; i < 3; i++) {
        header2.push('fatturato', 'bgt', 'fatt/ bgt');
      }
      header2.push('fatturato', 'bgt', 'fatt/ bgt', 'fatt / bgt');

      const row2 = worksheet.addRow(header2);
      row2.eachCell((cell, colNumber) => {
        cell.border = BORDER_STYLE;
        if (colNumber > 1) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }; // Light Blue
          cell.alignment = { horizontal: 'center' };
          cell.font = { size: 10 };
        }
      });

      currentRow++;

      // --- Data Rows ---
      budgetSituationData.forEach(person => {
        const rowData: any[] = [person.name];
        
        let qActual = 0;
        let qBudget = 0;

        monthsInQ.forEach(mIdx => {
          const m = person.months[mIdx];
          rowData.push(m.actual || 0, m.budget || 0, (m.actual || 0) - (m.budget || 0));
          qActual += (m.actual || 0);
          qBudget += (m.budget || 0);
        });

        // Quarter Stats
        const qDiff = qActual - qBudget;
        const qPerc = qBudget > 0 ? (qActual / qBudget - 1) : -1;

        rowData.push(qActual, qBudget, qDiff, qPerc);

        const row = worksheet.addRow(rowData);
        row.eachCell((cell, colNumber) => {
          cell.border = BORDER_STYLE;
          if (colNumber === 1) {
            cell.font = { size: 11 };
          } else if (colNumber === 14) {
            // Percentage column
            cell.numFmt = '0%';
            cell.font = { bold: true, color: { argb: (cell.value as number) >= 0 ? 'FF008000' : 'FFFF0000' } };
            cell.alignment = { horizontal: 'center' };
          } else {
            cell.numFmt = '#,##0';
          }
        });
        currentRow++;
      });

      // --- Totals Row ---
      const totalsData: any[] = [''];
      let grandQActual = 0;
      let grandQBudget = 0;

      monthsInQ.forEach(mIdx => {
        const mActual = budgetSituationData.reduce((acc, p) => acc + (p.months[mIdx].actual || 0), 0);
        const mBudget = budgetSituationData.reduce((acc, p) => acc + (p.months[mIdx].budget || 0), 0);
        totalsData.push(mActual, mBudget, mActual - mBudget);
        grandQActual += mActual;
        grandQBudget += mBudget;
      });

      const grandQDiff = grandQActual - grandQBudget;
      const grandQPerc = grandQBudget > 0 ? (grandQActual / grandQBudget - 1) : -1;
      totalsData.push(grandQActual, grandQBudget, grandQDiff, grandQPerc);

      const totRow = worksheet.addRow(totalsData);
      totRow.eachCell((cell, colNumber) => {
        cell.border = BORDER_STYLE;
        cell.font = { bold: true };
        if (colNumber > 1 && colNumber < 14) {
          cell.numFmt = '#,##0';
        } else if (colNumber === 14) {
          cell.numFmt = '0%';
          cell.font = { bold: true, color: { argb: (cell.value as number) >= 0 ? 'FF008000' : 'FFFF0000' } };
          cell.alignment = { horizontal: 'center' };
        }
      });

      currentRow++;
      worksheet.addRow([]); // Gap between blocks
      currentRow++;
    }

    // Column Widths
    worksheet.getColumn(1).width = 30;
    for (let i = 2; i <= 14; i++) {
      worksheet.getColumn(i).width = 12;
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Situazione_Budget_${selectedYear}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const generateAndDownloadExcel = async (patients: any[], filename: string, split: boolean = false) => {
    if (patients.length === 0) return;
    const workbook = new ExcelJS.Workbook();
    
    // Determine column structure based on filename or content if needed
    // But better use a flag or check if it's the specific verification
    const isContattiChiamati = filename.includes('Verifica_contattichiamati');
    const isMapSelection = filename.includes('SelezioneMappa');

    let columns = [
      { header: 'TIPO', key: 'tipo', width: 15 },
      { header: 'CLASSIFICAZIONE', key: 'classification', width: 20 },
      { header: 'COGNOME', key: 'cognome', width: 20 },
      { header: 'NOME', key: 'nome', width: 20 },
      { header: 'INDIRIZZO', key: 'indirizzo', width: 30 },
      { header: 'CITTÀ', key: 'citta', width: 20 },
      { header: 'PROVINCIA', key: 'provincia', width: 10 },
      { header: 'CAP', key: 'cap', width: 10 },
      { header: 'TELEFONO', key: 'telefono', width: 15 },
      { header: 'ID PAZIENTE', key: 'id', width: 15 },
      { header: 'FILIALE', key: 'store', width: 20 },
      { header: 'ULTIMA TELEF', key: 'lastCall', width: 15 },
      { header: 'ULTIMO APP', key: 'lastApp', width: 15 },
      { header: 'MOTIVI PRIORITÀ', key: 'reasons', width: 40 },
      { header: 'NOTE', key: 'note', width: 50 }
    ];

    if (isMapSelection) {
      columns = [
        { header: 'TIPO', key: 'tipo', width: 15 },
        { header: 'COGNOME', key: 'cognome', width: 20 },
        { header: 'NOME', key: 'nome', width: 20 },
        { header: 'TELEFONO', key: 'telefono', width: 15 },
        { header: 'INDIRIZZO', key: 'indirizzo', width: 30 },
        { header: 'CITTÀ', key: 'citta', width: 20 },
        { header: 'CAP', key: 'cap', width: 10 },
        { header: 'PROVINCIA', key: 'provincia', width: 10 },
        { header: 'STORE', key: 'store', width: 20 },
        { header: 'ULTIMA TELEF', key: 'lastCall', width: 15 },
        { header: 'ULTIMO APP', key: 'lastApp', width: 15 },
        { header: 'MOTIVI INTERESSE', key: 'reasons', width: 40 }
      ];
    } else if (isContattiChiamati) {
      columns = [
        { header: 'TIPO', key: 'tipo', width: 15 },
        { header: 'COGNOME', key: 'cognome', width: 20 },
        { header: 'NOME', key: 'nome', width: 20 },
        { header: 'TELEFONO', key: 'telefono', width: 15 },
        { header: 'INDIRIZZO', key: 'indirizzo', width: 30 },
        { header: 'CITTÀ', key: 'citta', width: 20 },
        { header: 'CAP', key: 'cap', width: 10 },
        { header: 'PROVINCIA', key: 'provincia', width: 10 },
        { header: 'FILIALE', key: 'store', width: 20 },
        { header: 'ULTIMA TELEF', key: 'lastCall', width: 15 },
        { header: 'ULTIMO APP', key: 'lastApp', width: 15 }
      ];
    }

    const applyStyling = (sheet: ExcelJS.Worksheet) => {
      const headerRow = sheet.getRow(1);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
      headerRow.alignment = { vertical: 'middle', horizontal: 'left' };
      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
            left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
            bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
            right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
          };
        });
        if (rowNumber % 2 === 0) {
          row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
        }
      });
    };

    const formatPatientRow = (p: any) => ({
      tipo: p.tipo || '',
      classification: p.classification || '',
      classificationReason: p.classificationReason || '',
      cognome: p.cognome || 'N/D',
      nome: p.nome || 'N/D',
      indirizzo: p.indirizzo || 'N/D',
      citta: p.citta || 'N/D',
      provincia: p.provincia || '',
      cap: p.cap || 'N/D',
      telefono: p.telefono || 'N/D',
      id: p.id || 'N/D',
      store: p.store || 'N/D',
      lastCall: p.lastCall || 'N/D',
      lastApp: p.lastApp || 'N/D',
      reasons: Array.isArray(p.reasons) ? p.reasons.join(', ') : '',
      note: p.note || ''
    });

    if (split) {
      const freddeData = patients.filter(p => {
        const tipo = (p.tipo || '').trim().toUpperCase();
        return tipo === 'LEAD' || tipo === '' || tipo === 'CONTATTO CHIAMATO';
      });

      const pcData = patients.filter(p => {
        const tipo = (p.tipo || '').trim().toUpperCase();
        return tipo === 'PC' || tipo === 'PREMATURO' || tipo === 'CLIENTE CONCORRENZA';
      });

      // Create "Fredde" sheet
      const freddeSheet = workbook.addWorksheet('Fredde');
      freddeSheet.columns = columns;
      freddeData.forEach(p => freddeSheet.addRow(formatPatientRow(p)));
      applyStyling(freddeSheet);

      // Create "PC" sheet
      const pcSheet = workbook.addWorksheet('PC');
      pcSheet.columns = columns;
      pcData.forEach(p => pcSheet.addRow(formatPatientRow(p)));
      applyStyling(pcSheet);
    } else {
      const sheet = workbook.addWorksheet('Dati');
      sheet.columns = columns;
      patients.forEach(p => sheet.addRow(formatPatientRow(p)));
      applyStyling(sheet);
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const downloadOpenTrialsExcel = async () => {
    if (openTrialsAnalysis.length === 0) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Prove Aperte');

    sheet.columns = [
      { header: 'Audioprotesista', key: 'audioprotesista', width: 25 },
      { header: 'Cliente finale', key: 'cliente', width: 30 },
      { header: 'Canale', key: 'canale', width: 25 },
      { header: 'Data Documento', key: 'fullDate', width: 15 },
      { header: 'Giorni Trascorsi', key: 'daysElapsed', width: 15 },
      { header: 'Valore Totale', key: 'valore', width: 15 }
    ];

    openTrialsAnalysis.forEach(([audioprotesista, trials]) => {
      trials.forEach((t: any) => {
        sheet.addRow({
          audioprotesista,
          cliente: t.cliente || 'N/D',
          canale: t.canale || 'N/D',
          fullDate: t.fullDate || 'N/D',
          daysElapsed: t.daysElapsed,
          valore: t.valore || 0
        });
      });
      // Add summary row with formatting
      const summaryRow = sheet.addRow({
        audioprotesista: `TOTALE ${audioprotesista}`,
        valore: trials.reduce((sum: number, t: any) => sum + (t.valore || 0), 0)
      });
      summaryRow.font = { bold: true };
      summaryRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      summaryRow.getCell('valore').numFmt = '#,##0.00 "€"';
      sheet.addRow({}); 
    });

    // Add Grand Total
    const grandTotal = openTrialsAnalysis.reduce((sum, [_, trials]) => 
      sum + trials.reduce((s, t: any) => s + (t.valore || 0), 0), 0
    );
    const grandRow = sheet.addRow({
      audioprotesista: 'TOTALE COMPLESSIVO',
      valore: grandTotal
    });
    grandRow.font = { bold: true, size: 12 };
    grandRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow
    grandRow.getCell('valore').numFmt = '#,##0.00 "€"';

    // Format all value cells
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const valueCell = row.getCell('valore');
        if (valueCell.value && typeof valueCell.value === 'number') {
          valueCell.numFmt = '#,##0.00 "€"';
        }
      }
    });

    // Style header
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Prove_Aperte_${new Date().toLocaleDateString('it-IT').replace(/\//g, '-')}.xlsx`;
    anchor.click();
    window.URL.revokeObjectURL(url);
  };

  const getNoContattatiPatients = (branch: string, storeId: string, months: number, patientType: string, cap: string) => {
    const now = new Date();
    const thresholdTimestamp = now.getTime() - (months * 30 * 24 * 60 * 60 * 1000);
    const twoMonthsAgo = now.getTime() - (60 * 24 * 60 * 60 * 1000);

    const filteredPatients = patientsData.filter(p => {
      const isCorrectType = !['Deceduto', 'Cliente', 'Errore', 'Errato', 'Normoudente'].includes(p.tipo);
      if (!isCorrectType) return false;

      let matches = true;
      if (branch) matches = matches && String(p.store || '').toLowerCase().trim() === branch.toLowerCase().trim();
      if (patientType) matches = matches && String(p.tipo || '').toLowerCase().trim() === patientType.toLowerCase().trim();
      if (cap) matches = matches && String(p.cap || '').trim().includes(cap.trim());
      if (storeId) {
        const store = storesData.find(s => s.id === storeId);
        if (store) {
          const associatedCaps = storeCaps.filter(sc => sc.storeId === storeId).map(sc => sc.cap);
          const allStoreCaps = [store.cap, ...associatedCaps].filter(Boolean);
          if (allStoreCaps.length > 0) matches = matches && allStoreCaps.includes(p.cap);
        }
      }
      return matches;
    });

    const normalizeId = (id: any): string | undefined => {
      if (!id) return undefined;
      const s = String(id).trim();
      if (/^\d+$/.test(s)) return String(parseInt(s, 10));
      return s;
    };

    const callsById = new Map<string, any[]>();
    const callsByName = new Map<string, any[]>();
    callsData.forEach(c => {
      const id = normalizeId(c.pazienteId);
      if (id) {
        if (!callsById.has(id)) callsById.set(id, []);
        callsById.get(id)!.push(c);
      }
      if (c.nome) {
        const name = c.nome.toLowerCase().trim();
        if (!callsByName.has(name)) callsByName.set(name, []);
        callsByName.get(name)!.push(c);
      }
    });

    const appsById = new Map<string, any[]>();
    const appsByName = new Map<string, any[]>();
    appointmentsData.forEach(a => {
      const id = normalizeId(a.pazienteId);
      if (id) {
        if (!appsById.has(id)) appsById.set(id, []);
        appsById.get(id)!.push(a);
      }
      if (a.nome) {
        const name = a.nome.toLowerCase().trim();
        if (!appsByName.has(name)) appsByName.set(name, []);
        appsByName.get(name)!.push(a);
      }
    });

    return filteredPatients.filter(p => {
      const pId = normalizeId(p.id);
      const pName = p.nomeCompleto ? p.nomeCompleto.toLowerCase().trim() : null;
      const calls = [...(pId ? (callsById.get(pId) || []) : []), ...(pName ? (callsByName.get(pName) || []) : [])];
      const apps = [...(pId ? (appsById.get(pId) || []) : []), ...(pName ? (appsByName.get(pName) || []) : [])];

      const hasFutureApp = apps.some(a => (a.esito || '').toLowerCase() === 'appuntamento futuro');
      const hasRecentAppInTwoMonths = apps.some(a => a.timestamp >= twoMonthsAgo && a.timestamp <= now.getTime());
      if (hasFutureApp || hasRecentAppInTwoMonths) return false;

      const allNotes = [p.note, ...calls.map(c => c.note), ...apps.map(a => a.note)].join(' ').toLowerCase();
      if (allNotes.includes('non interessato')) return false;

      const hasRecentCall = calls.some(c => c.timestamp && c.timestamp >= thresholdTimestamp);
      const hasRecentApp = apps.some(a => a.timestamp && a.timestamp >= thresholdTimestamp);
      return !hasRecentCall && !hasRecentApp;
    }).map(p => {
      const pId = normalizeId(p.id);
      const pName = p.nomeCompleto ? p.nomeCompleto.toLowerCase().trim() : null;
      const calls = [...(pId ? (callsById.get(pId) || []) : []), ...(pName ? (callsByName.get(pName) || []) : [])];
      const apps = [...(pId ? (appsById.get(pId) || []) : []), ...(pName ? (appsByName.get(pName) || []) : [])];
      const lastCallTimestamp = calls.reduce((max: number, c) => (c.timestamp && c.timestamp > max) ? c.timestamp : max, 0);
      const lastAppTimestamp = apps.reduce((max: number, a) => (a.timestamp && a.timestamp > max) ? a.timestamp : max, 0);
      return {
        ...p,
        score: 0,
        reasons: [`Non contattato da ${months} mesi`],
        lastCall: lastCallTimestamp ? new Date(lastCallTimestamp).toLocaleDateString('it-IT') : 'Mai',
        lastApp: lastAppTimestamp ? new Date(lastAppTimestamp).toLocaleDateString('it-IT') : 'Mai'
      };
    });
  };

  const handleDownloadResults = async () => {
    if (analyzedPatients.length === 0) return;
    const selectedStore = storesData.find(s => s.id === selectedStoreId);
    const storeName = selectedStore ? selectedStore.nome : selectedBranch;
    const filename = `${storeName}-SelezioneMappa.xlsx`;
    await generateAndDownloadExcel(analyzedPatients, filename, false);
  };

  return (
    <div className="min-h-screen bg-[#E4E3E0] text-[#141414] font-sans p-6 md:p-12">
      <header className="mb-12 border-b border-[#141414] pb-8">
        <div className="flex flex-col md:flex-row md:items-end justify-between gap-6">
          <div>
            <h1 className="text-5xl font-bold tracking-tighter mb-4">AskoltaOra</h1>
            <div className="flex gap-6 mb-4">
              <button 
                onClick={() => setActiveTab('sales')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'sales' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Analisi Vendite
              </button>
              <button 
                onClick={() => setActiveTab('openTrials')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'openTrials' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Prove Aperte
              </button>
              <button 
                onClick={() => setActiveTab('patients')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'patients' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Gestione Pazienti
              </button>
              <button 
                onClick={() => setActiveTab('appointmentsManagement')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'appointmentsManagement' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Gestione Appuntamenti
              </button>
              <button 
                onClick={() => setActiveTab('screeningEntrances')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'screeningEntrances' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                ingressi Screening
              </button>
              <button 
                onClick={() => setActiveTab('noah')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'noah' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Analizza Noah
              </button>
              <button 
                onClick={() => setActiveTab('comparisons')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'comparisons' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Comparazioni
              </button>
              <button 
                onClick={() => setActiveTab('cap')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'cap' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Carica CAP
              </button>
              <button 
                onClick={() => setActiveTab('recapiti')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'recapiti' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Pazienti-Recapiti
              </button>
              <button 
                onClick={() => setActiveTab('verifiche')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'verifiche' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Verifiche
              </button>
              <button 
                onClick={() => setActiveTab('channels')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'channels' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Analizza Canali
              </button>
              <button 
                onClick={() => setActiveTab('audioAnalysis')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'audioAnalysis' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Analisi Audio
              </button>
              <button 
                onClick={() => setActiveTab('budget')}
                className={cn(
                  "text-[10px] uppercase tracking-[0.3em] pb-1 border-b-2 transition-all",
                  activeTab === 'budget' ? "border-[#141414] opacity-100" : "border-transparent opacity-40 hover:opacity-60"
                )}
              >
                Budget
              </button>
            </div>
            <div className="flex items-center gap-4">
              <p className="font-serif italic text-lg opacity-60">CRM {
                activeTab === 'sales' ? 'Sales Dashboard' : 
                activeTab === 'openTrials' ? 'Prove Aperte' :
                activeTab === 'patients' ? 'Gestione Pazienti' : 
                activeTab === 'noah' ? 'Analisi Noah' : 
                activeTab === 'cap' ? 'Strutture Esterne' :
                activeTab === 'recapiti' ? 'Pazienti-Recapiti' :
                activeTab === 'verifiche' ? 'Verifiche Anagrafiche' :
                activeTab === 'channels' ? 'Analisi Canali' :
                activeTab === 'audioAnalysis' ? 'Analisi Audio' :
                activeTab === 'budget' ? 'Gestione Budget' :
                'Comparazioni Annuali'
              }</p>
              {(activeTab === 'sales' || activeTab === 'channels' || activeTab === 'budget') && (
                <select 
                  value={selectedYear}
                  onChange={(e) => setSelectedYear(Number(e.target.value))}
                  className="bg-transparent border-b border-[#141414] font-mono font-bold text-lg focus:outline-none cursor-pointer"
                >
                  {availableYears.map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              )}
            </div>
          </div>
          {activeTab === 'sales' && (
            <div className="flex gap-4">
              <div className="bg-white border border-[#141414] p-4 min-w-[160px]">
                <p className="text-[10px] uppercase tracking-widest opacity-50 mb-1">Totale Netto {selectedYear}</p>
                <p className="text-2xl font-mono font-bold">{totalSales.toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 0, maximumFractionDigits: 0 })}</p>
              </div>
            </div>
          )}
          {activeTab === 'budget' && (
            <div className="flex gap-4">
              <div className="bg-[#141414] text-white p-4 min-w-[160px]">
                <p className="text-[10px] uppercase tracking-widest opacity-50 mb-1">Totale Budget Annuale</p>
                <p className="text-2xl font-mono font-bold text-emerald-400">{personnelBudgets.reduce((acc, b) => acc + b.annualBudget, 0).toLocaleString('it-IT')}</p>
              </div>
            </div>
          )}
        </div>
      </header>

      {activeTab === 'sales' ? (
        <div className="flex flex-col gap-4 mb-8">
          {/* Upload Section */}
          <div className="w-full grid grid-cols-1 lg:grid-cols-4 gap-4 mb-2">
            <div className="lg:col-span-3 bg-white border border-[#141414] p-3 flex flex-col justify-center">
              <div className="flex items-center justify-between mb-2">
                <h2 className="font-serif italic text-lg flex items-center gap-2">
                  <Upload size={18} /> Carica Dati
                </h2>
                <p className="text-[10px] opacity-50">Seleziona i file Excel dalla cartella locale.</p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-4 gap-3 items-center">
                <label className="md:col-span-1 flex flex-col items-center justify-center h-14 border-2 border-dashed border-gray-200 hover:border-[#141414] transition-colors cursor-pointer group bg-gray-50/50">
                  <input 
                    type="file" 
                    accept=".xlsx, .xls" 
                    multiple
                    className="hidden" 
                    onChange={handleMultipleSalesUpload}
                  />
                  <Upload size={14} className="text-gray-400 group-hover:text-[#141414] mb-0.5" />
                  <span className="text-[8px] font-bold uppercase tracking-wider">CARICA FILE</span>
                </label>

                <div className="md:col-span-3 flex gap-2">
                  {[
                    { label: 'VENDITE', loaded: salesData.length > 0, name: 'ff.xlsx' },
                    { label: 'STORNI', loaded: creditNotes.length > 0, name: 'rr.xlsx' },
                    { label: 'PROVE', loaded: trialsData.length > 0, name: 'pp.xlsx' }
                  ].map((f) => (
                    <div 
                      key={f.label}
                      className={cn(
                        "flex-1 p-2 border flex flex-col justify-center",
                        f.loaded ? "bg-emerald-50 border-emerald-200" : "bg-gray-50 border-gray-100 opacity-60"
                      )}
                    >
                      <div className="flex justify-between items-center mb-0.5">
                        <span className="text-[8px] font-bold tracking-widest">{f.label}</span>
                        {f.loaded && <CheckCircle2 size={10} className="text-emerald-500" />}
                      </div>
                      <p className="text-[9px] font-mono truncate">{f.loaded ? `${f.name} Caricato` : `${f.name} Mancante`}</p>
                    </div>
                  ))}
                </div>
              </div>

              {error && (
                <div className="mt-2 p-2 bg-red-50 border border-red-200 text-red-700 text-[10px] flex items-start gap-2">
                  <AlertCircle size={10} className="mt-0.5 shrink-0" />
                  <span>{error}</span>
                </div>
              )}
            </div>

            <div className="lg:col-span-1 bg-[#141414] text-white p-3 flex flex-col justify-center">
              <h3 className="text-[9px] uppercase tracking-[0.2em] mb-2 opacity-60">Filtri Attivi</h3>
              <ul className="text-[9px] space-y-1 font-mono opacity-80 overflow-y-auto max-h-20">
                <li className="flex gap-1.5 text-red-400 items-start">
                  <AlertCircle size={10} className="mt-0.5 shrink-0" />
                  <span>Escluso: Askoltaora Cuneo</span>
                </li>
                <li className="flex gap-1.5 text-red-400 items-start">
                  <AlertCircle size={10} className="mt-0.5 shrink-0" />
                  <span>Escluso: Askoltaora Mondovì</span>
                </li>
                <li className="flex gap-1.5 text-emerald-400 items-start">
                  <CheckCircle2 size={10} className="mt-0.5 shrink-0" />
                  <span>Importi ≥ 600€ (ff.xlsx)</span>
                </li>
              </ul>
            </div>
          </div>

          {/* Table Section */}
          <div className="w-full">
            <div className="bg-white border border-[#141414] overflow-hidden">
              <div className="p-6 border-b border-[#141414] flex justify-between items-center">
                <h2 className="font-serif italic text-xl">Riepilogo Vendite {selectedYear}</h2>
                <div className="flex gap-4">
                  <button
                    onClick={() => setShowAllDetails(!showAllDetails)}
                    disabled={salesData.length === 0}
                    className={cn(
                      "px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                      salesData.length === 0 ? "opacity-30 cursor-not-allowed" : "hover:bg-gray-50 active:scale-95",
                      showAllDetails ? "bg-[#141414] text-white" : "bg-white text-[#141414]"
                    )}
                  >
                    <ChevronRight size={14} className={cn("transition-transform", showAllDetails && "rotate-90")} />
                    {showAllDetails ? "Nascondi Canali" : "Mostra Canali"}
                  </button>
                  <button
                    onClick={handleMarketingAnalysis}
                    disabled={salesData.length === 0}
                    className={cn(
                      "px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                      salesData.length === 0 ? "opacity-30 cursor-not-allowed" : "hover:bg-gray-50 active:scale-95"
                    )}
                  >
                    <TrendingUp size={14} /> Analisi Canali Marketing
                  </button>
                  <button
                    onClick={handleScreeningAnalysis}
                    disabled={salesData.length === 0}
                    className={cn(
                      "px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                      salesData.length === 0 ? "opacity-30 cursor-not-allowed" : "hover:bg-gray-50 active:scale-95"
                    )}
                  >
                    <Users size={14} /> Analisi Screening
                  </button>
                  <button
                    onClick={handleCoefficientiAnalysis}
                    disabled={salesData.length === 0}
                    className={cn(
                      "px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                      salesData.length === 0 ? "opacity-30 cursor-not-allowed" : "hover:bg-gray-50 active:scale-95"
                    )}
                  >
                    <BarChart3 size={14} /> Coefficienti
                  </button>
                  <button
                    onClick={handleMediciAnalysis}
                    disabled={salesData.length === 0}
                    className={cn(
                      "px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                      salesData.length === 0 ? "opacity-30 cursor-not-allowed" : "hover:bg-gray-50 active:scale-95"
                    )}
                  >
                    <Stethoscope size={14} /> Medici
                  </button>
                  <button
                    onClick={handleDownloadSalesExcel}
                    disabled={salesData.length === 0}
                    className={cn(
                      "px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2 bg-emerald-600 text-white",
                      salesData.length === 0 ? "opacity-30 cursor-not-allowed" : "hover:bg-emerald-700 active:scale-95"
                    )}
                  >
                    <Download size={14} /> Download Excel
                  </button>
                  <div className="flex items-center gap-2 text-[10px] uppercase tracking-widest opacity-60">
                    <div className="w-2 h-2 bg-emerald-500"></div> Vendite
                  </div>
                  <div className="flex items-center gap-2 text-[10px] uppercase tracking-widest opacity-60">
                    <div className="w-2 h-2 bg-red-400"></div> Storni
                  </div>
                </div>
              </div>
              
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-gray-50 border-b border-[#141414]">
                      <th className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 border-r border-[#141414]">Audioprotesista</th>
                      <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-[#141414] text-center bg-gray-100/50">Totale Anno</th>
                      <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-[#141414] text-center bg-emerald-50/50 text-emerald-700">Budget Progr.</th>
                      {MONTHS.map(m => (
                        <th key={m} className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 text-center min-w-[120px]">
                          {m.slice(0, 3)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="font-mono text-[11px]">
                    {aggregatedData.length > 0 ? (
                      <>
                        {aggregatedData.map(([name, data], idx) => {
                          const rowNetTotal = data.sales.reduce((a, b) => a + b, 0) - data.credits.reduce((a, b) => a + b, 0);
                          const personBudget = personnelBudgets.find(b => b.name === name);
                          const annualBudget = personBudget?.annualBudget || 0;
                          
                          const currentYear = new Date().getFullYear();
                          const currentMonth = new Date().getMonth();
                          const progMonthIdx = selectedYear < currentYear ? 11 : selectedYear > currentYear ? -1 : currentMonth;
                          const progPct = monthlyDistribution.slice(0, progMonthIdx + 1).reduce((a, b) => a + b, 0);
                          const budgetProgr = (annualBudget * progPct) / 100;
                          
                          return (
                            <React.Fragment key={name}>
                              <tr className={cn(
                                "border-b border-[#141414] hover:bg-gray-50 transition-colors",
                                idx % 2 === 0 ? "bg-white" : "bg-gray-50/30"
                              )}>
                                <td className="p-4 font-bold border-r border-[#141414] whitespace-nowrap">
                                  {name}
                                </td>
                                <td 
                                  onClick={() => handleCellClick(name, -1)}
                                  className="p-4 text-center font-bold border-r border-[#141414] bg-gray-50/30 cursor-pointer hover:bg-gray-100 transition-colors"
                                >
                                  {rowNetTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                </td>
                                <td className="p-4 text-center font-bold border-r border-[#141414] bg-emerald-50/20 text-emerald-700">
                                  {budgetProgr.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                  <div className="text-[9px] opacity-40 font-normal">
                                    {(rowNetTotal - budgetProgr) >= 0 ? '+' : ''}{(rowNetTotal - budgetProgr).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                  </div>
                                </td>
                                {data.sales.map((saleVal, mIdx) => {
                                  const creditVal = data.credits[mIdx];
                                  const net = saleVal - creditVal;
                                  return (
                                    <td 
                                      key={mIdx} 
                                      onClick={() => handleCellClick(name, mIdx)}
                                      className={cn(
                                        "p-4 text-center border-r border-gray-100 last:border-r-0 cursor-pointer transition-colors",
                                        net < 0 ? "text-red-500 hover:bg-red-50" : net > 0 ? "text-[#141414] hover:bg-gray-50" : "opacity-20 hover:bg-gray-50"
                                      )}
                                    >
                                      <div className="font-mono font-bold">
                                        {net !== 0 ? net.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 }) : '-'}
                                      </div>
                                    </td>
                                  );
                                })}
                              </tr>
                              {showAllDetails && (
                                <>
                                  {(Object.entries(data.channelMonthly) as [string, number[]][])
                                    .sort((a, b) => b[1].reduce((sum, v) => sum + v, 0) - a[1].reduce((sum, v) => sum + v, 0))
                                    .map(([channel, monthlyValues]) => {
                                      const channelTotal = monthlyValues.reduce((sum, v) => sum + v, 0);
                                      if (channelTotal === 0) return null;
                                      
                                      return (
                                        <tr key={`${name}-${channel}`} className="bg-gray-50/50 border-b border-gray-100 text-[10px]">
                                          <td className="p-3 pl-8 italic opacity-60 border-r border-[#141414] whitespace-nowrap">
                                            ↳ {channel}
                                          </td>
                                          <td 
                                            onClick={() => handleCellClick(name, -1, channel)}
                                            className="p-3 text-center font-bold border-r border-[#141414] opacity-60 bg-gray-50/30 cursor-pointer hover:bg-gray-100 transition-colors"
                                          >
                                            {channelTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                          </td>
                                          <td className="p-3 border-r border-[#141414] bg-emerald-50/10 opacity-30"></td>
                                          {monthlyValues.map((val, mIdx) => (
                                            <td 
                                              key={mIdx} 
                                              onClick={() => handleCellClick(name, mIdx, channel)}
                                              className={cn(
                                                "p-3 text-center border-r border-gray-100 last:border-r-0 opacity-60 transition-colors",
                                                val !== 0 ? "cursor-pointer hover:bg-gray-100" : ""
                                              )}
                                            >
                                              {val !== 0 ? val.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 }) : '-'}
                                            </td>
                                          ))}
                                        </tr>
                                      );
                                    })}
                                </>
                              )}
                            </React.Fragment>
                          );
                        })}
                        {/* Totals Row */}
                        <tr className="bg-[#141414] text-white font-bold">
                          <td className="p-4 border-r border-white/20 uppercase tracking-widest text-[9px]">Totale Mensile (Netto)</td>
                          <td 
                            onClick={() => handleCellClick(null, -1)}
                            className="p-4 text-center border-r border-white/20 bg-white/10 cursor-pointer hover:bg-white/20 transition-colors"
                          >
                            {totalSales.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                          </td>
                          <td className="p-4 text-center border-r border-white/20 bg-white/5 text-emerald-400">
                            {(() => {
                              const currentYear = new Date().getFullYear();
                              const currentMonth = new Date().getMonth();
                              const progMonthIdx = selectedYear < currentYear ? 11 : selectedYear > currentYear ? -1 : currentMonth;
                              const progPct = monthlyDistribution.slice(0, progMonthIdx + 1).reduce((a, b) => a + b, 0);
                              const totalAnnualBudget = personnelBudgets.reduce((acc, b) => acc + b.annualBudget, 0);
                              const totalBudgetProgr = (totalAnnualBudget * progPct) / 100;
                              return (
                                <>
                                  {totalBudgetProgr.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                  <div className="text-[9px] opacity-60 font-normal">
                                    {(totalSales - totalBudgetProgr) >= 0 ? '+' : ''}{(totalSales - totalBudgetProgr).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                  </div>
                                </>
                              );
                            })()}
                          </td>
                          {Array.from({ length: 12 }).map((_, mIdx) => {
                            const monthSales = aggregatedData.reduce((acc, [_, d]) => acc + d.sales[mIdx], 0);
                            const monthCredits = aggregatedData.reduce((acc, [_, d]) => acc + d.credits[mIdx], 0);
                            const monthNet = monthSales - monthCredits;
                            return (
                              <td 
                                key={mIdx} 
                                onClick={() => monthNet !== 0 && handleCellClick(null, mIdx)}
                                className={cn(
                                  "p-4 text-center border-r border-white/10 last:border-r-0 transition-colors",
                                  monthNet !== 0 ? "cursor-pointer hover:bg-white/10" : ""
                                )}
                              >
                                {monthNet !== 0 ? monthNet.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 }) : '-'}
                              </td>
                            );
                          })}
                        </tr>
                      </>
                    ) : (
                      <tr>
                        <td colSpan={15} className="p-12 text-center opacity-40 italic font-serif text-lg">
                          Carica i file ff.xlsx e rr.xlsx per visualizzare i dati
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      ) : activeTab === 'openTrials' ? (
        <div className="space-y-6">
          <div className="bg-white border border-[#141414] p-10">
            <div className="flex items-center justify-between mb-10">
              <div className="flex flex-col gap-1">
                <h3 className="text-xl font-serif italic flex items-center gap-3">
                  <Stethoscope size={24} />
                  Elenco Prove Aperte per Audioprotesista
                </h3>
                {openTrialsAnalysis.length > 0 && (
                  <div className="text-sm font-bold bg-[#141414] text-white px-4 py-1 self-start mt-2">
                    TOTALE COMPLESSIVO: € {openTrialsAnalysis.reduce((sum, [_, trials]) => 
                      sum + trials.reduce((s: number, t: any) => s + (t.valore || 0), 0), 0
                    ).toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                )}
              </div>
              <button 
                onClick={downloadOpenTrialsExcel}
                className="flex items-center gap-2 px-4 py-2 border border-[#141414] text-[10px] uppercase tracking-widest hover:bg-[#141414] hover:text-white transition-all disabled:opacity-30 disabled:pointer-events-none"
                disabled={openTrialsAnalysis.length === 0}
              >
                <Download size={14} /> Download Excel
              </button>
            </div>
            
            {openTrialsAnalysis.length === 0 ? (
              <div className="py-20 text-center opacity-40 italic font-serif">Nessuna prova aperta trovata. Carica il file pp.xlsx per visualizzare i dati.</div>
            ) : (
              <div className="space-y-16">
                {openTrialsAnalysis.map(([audioprotesista, trials]) => (
                  <div key={audioprotesista}>
                    <div className="flex items-center justify-between mb-6 pb-2 border-b border-[#141414]/10">
                      <div className="flex flex-col">
                        <span className="font-bold text-sm uppercase tracking-[0.2em]">{audioprotesista}</span>
                        <span className="text-[10px] opacity-40 uppercase tracking-widest mt-1">
                          Totale: € {trials.reduce((sum: number, t: any) => sum + (t.valore || 0), 0).toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                        </span>
                      </div>
                      <span className="text-[10px] opacity-40 uppercase tracking-widest">{trials.length} prove aperte</span>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left font-mono text-[11px] border-collapse">
                        <thead>
                          <tr className="border-b border-[#141414] opacity-50">
                            <th className="py-4 px-2">Cliente finale</th>
                            <th className="py-4 px-2">Canale</th>
                            <th className="py-4 px-2">Data Documento</th>
                            <th className="py-4 px-2 text-center">Giorni trascorsi</th>
                            <th className="py-4 px-2 text-right">Valore totale</th>
                          </tr>
                        </thead>
                        <tbody>
                          {trials.map((t, idx) => (
                            <tr key={idx} className="border-b border-[#141414]/5 last:border-0 hover:bg-[#141414]/5 transition-colors">
                              <td className="py-4 px-2 font-bold">{t.cliente || 'N/D'}</td>
                              <td className="py-4 px-2 opacity-60">{t.canale || 'N/D'}</td>
                              <td className="py-4 px-2 opacity-60">{t.fullDate || 'N/D'}</td>
                              <td className="py-4 px-2 text-center">
                                <span className={cn(
                                  "px-2 py-1 font-bold",
                                  t.daysElapsed > 60 ? "bg-red-100 text-red-800" : 
                                  t.daysElapsed > 30 ? "bg-amber-100 text-amber-800" : 
                                  "bg-emerald-100 text-emerald-800"
                                )}>
                                  {t.daysElapsed} GG
                                </span>
                              </td>
                              <td className="py-4 px-2 text-right font-bold text-sm">
                                € {(t.valore || 0).toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="border-t border-[#141414] bg-[#141414]/5">
                            <td colSpan={4} className="py-4 px-2 text-right font-bold uppercase tracking-wider">Totale {audioprotesista}</td>
                            <td className="py-4 px-2 text-right font-bold text-sm">
                              € {trials.reduce((sum: number, t: any) => sum + (t.valore || 0), 0).toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      ) : activeTab === 'patients' ? (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 mb-12">
          {/* Sidebar / Upload Section Patients */}
          <div className="lg:col-span-1 space-y-6">
            <div className="bg-white border border-[#141414] p-6">
              <h2 className="font-serif italic text-xl mb-4 flex items-center gap-2">
                <Upload size={20} /> Carica Dati Pazienti
              </h2>
              
              <div className="space-y-6">
                <label className="block group cursor-pointer">
                  <div className="bg-[#141414] text-white p-4 flex items-center justify-center gap-3 hover:bg-[#333] transition-all active:scale-95">
                    <Upload size={18} />
                    <span className="text-[10px] uppercase tracking-[0.2em] font-bold">Carica File (ii.xlsx / tt.xlsx / aa.xlsx / store.xlsx / dd.xlsx)</span>
                  </div>
                  <input 
                    type="file" 
                    accept=".xlsx, .xls" 
                    multiple
                    className="hidden" 
                    onChange={handleMultiplePatientsUpload}
                  />
                </label>

                <div className="grid grid-cols-1 gap-3">
                  <div className={cn(
                    "border border-[#141414] p-3 flex flex-col gap-2 transition-colors",
                    patientsData.length > 0 ? "bg-emerald-50 border-emerald-500" : "opacity-40"
                  )}>
                    <div className="flex items-center justify-between">
                      <span className="text-[9px] uppercase tracking-widest font-bold">Anagrafica</span>
                      {patientsData.length > 0 && <CheckCircle2 size={12} className="text-emerald-600" />}
                    </div>
                    <span className="text-[10px] font-mono">{patientsData.length > 0 ? `ii.xlsx Caricato (${patientsData.length} pazienti)` : "ii.xlsx Mancante"}</span>
                  </div>

                  <div className={cn(
                    "border border-[#141414] p-3 flex flex-col gap-2 transition-colors",
                    callsData.length > 0 ? "bg-emerald-50 border-emerald-500" : "opacity-40"
                  )}>
                    <div className="flex items-center justify-between">
                      <span className="text-[9px] uppercase tracking-widest font-bold">Telefonate</span>
                      {callsData.length > 0 && <CheckCircle2 size={12} className="text-emerald-600" />}
                    </div>
                    <span className="text-[10px] font-mono">{callsData.length > 0 ? `tt.xlsx Caricato (${callsData.length} telefonate)` : "tt.xlsx Mancante"}</span>
                  </div>

                  <div className={cn(
                    "border border-[#141414] p-3 flex flex-col gap-2 transition-colors",
                    appointmentsData.length > 0 ? "bg-emerald-50 border-emerald-500" : "opacity-40"
                  )}>
                    <div className="flex items-center justify-between">
                      <span className="text-[9px] uppercase tracking-widest font-bold">Appuntamenti</span>
                      {appointmentsData.length > 0 && <CheckCircle2 size={12} className="text-emerald-600" />}
                    </div>
                    <span className="text-[10px] font-mono">{appointmentsData.length > 0 ? `aa.xlsx Caricato (${appointmentsData.length} appuntamenti)` : "aa.xlsx Mancante"}</span>
                  </div>

                  <div className={cn(
                    "border border-[#141414] p-3 flex flex-col gap-2 transition-colors",
                    storesData.length > 0 ? "bg-emerald-50 border-emerald-500" : "opacity-40"
                  )}>
                    <div className="flex items-center justify-between">
                      <span className="text-[9px] uppercase tracking-widest font-bold">Strutture Esterne</span>
                      {storesData.length > 0 ? <CheckCircle2 size={12} className="text-emerald-600" /> : loading.stores && <Loader2 size={12} className="animate-spin" />}
                    </div>
                    <span className="text-[10px] font-mono">
                      {storesData.length > 0 ? `store.xlsx Caricato (${storesData.length} strutture)` : loading.stores ? "Caricamento dal database..." : "store.xlsx Opzionale"}
                    </span>
                  </div>

                  <div className={cn(
                    "border border-[#141414] p-3 flex flex-col gap-2 transition-colors",
                    availabilityData.length > 0 ? "bg-emerald-50 border-emerald-500" : "opacity-40"
                  )}>
                    <div className="flex items-center justify-between">
                      <span className="text-[9px] uppercase tracking-widest font-bold">Disponibilità</span>
                      {availabilityData.length > 0 && <CheckCircle2 size={12} className="text-emerald-600" />}
                    </div>
                    <span className="text-[10px] font-mono">{availabilityData.length > 0 ? `dd.xlsx Caricato (${availabilityData.length} eventi)` : "dd.xlsx Opzionale"}</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-[#141414] text-white p-6">
              <h3 className="text-[10px] uppercase tracking-[0.2em] mb-4 opacity-60">Logica Priorità</h3>
              <ul className="text-xs space-y-3 font-mono opacity-80">
                <li className="flex gap-2">
                  <TrendingUp size={12} className="shrink-0 mt-0.5 text-emerald-400" />
                  <span>Alta: Interesse Apparecchi Acustici</span>
                </li>
                <li className="flex gap-2">
                  <Calendar size={12} className="shrink-0 mt-0.5 text-blue-400" />
                  <span>Media: Ipoacusico / Visita Udito</span>
                </li>
              </ul>
            </div>
          </div>

          {/* Patients List Section */}
          <div className="lg:col-span-2">
            <div className="bg-white border border-[#141414] overflow-hidden">
              <div className="p-6 border-b border-[#141414] flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <h2 className="font-serif italic text-xl">Pazienti da Chiamare (Priorità)</h2>
                <div className="flex flex-wrap items-center gap-4 w-full md:w-auto">
                  <div className="flex-1 md:flex-none">
                    <select
                      value={selectedBranch}
                      onChange={(e) => setSelectedBranch(e.target.value)}
                      className="w-full bg-transparent border-b border-[#141414] font-mono font-bold text-xs focus:outline-none cursor-pointer py-1"
                    >
                      <option value="">Seleziona Filiale...</option>
                      {availableBranches.map(b => (
                        <option key={b} value={b}>{b}</option>
                      ))}
                    </select>
                  </div>

                  <div className="flex-1 md:flex-none">
                    <select
                      value={selectedPatientType}
                      onChange={(e) => setSelectedPatientType(e.target.value)}
                      className="w-full bg-transparent border-b border-[#141414] font-mono font-bold text-xs focus:outline-none cursor-pointer py-1"
                    >
                      <option value="">Seleziona Tipo...</option>
                      {availablePatientTypes.map(t => (
                        <option key={t} value={t}>{t}</option>
                      ))}
                    </select>
                  </div>

                  <div className="flex-1 md:flex-none relative">
                    <div className="flex flex-col">
                      <input 
                        type="text"
                        placeholder="Cerca CAP..."
                        value={capFilterSearch}
                        onChange={(e) => {
                          setCapFilterSearch(e.target.value);
                          setShowCapFilterDropdown(true);
                        }}
                        onFocus={() => setShowCapFilterDropdown(true)}
                        className="w-full bg-transparent border-b border-[#141414] font-mono font-bold text-xs focus:outline-none py-1"
                      />
                      {showCapFilterDropdown && (
                        <div className="absolute top-full left-0 w-full max-h-60 overflow-y-auto bg-white border border-[#141414] z-[100] shadow-xl mt-1">
                          <div 
                            className="p-2 hover:bg-gray-100 cursor-pointer text-[10px] uppercase tracking-widest font-bold border-b border-gray-100"
                            onClick={() => {
                              setSelectedCap('');
                              setCapFilterSearch('');
                              setShowCapFilterDropdown(false);
                            }}
                          >
                            Tutti i CAP
                          </div>
                          {availableCaps
                            .filter(c => c.includes(capFilterSearch))
                            .map(c => (
                              <div 
                                key={c}
                                className="p-2 hover:bg-gray-100 cursor-pointer text-[10px] font-mono border-b border-gray-100 last:border-0"
                                onClick={() => {
                                  setSelectedCap(c);
                                  setCapFilterSearch(c);
                                  setShowCapFilterDropdown(false);
                                }}
                              >
                                {c}
                              </div>
                            ))
                          }
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="flex-1 md:flex-none relative">
                    <div className="flex flex-col">
                      <input 
                        type="text"
                        placeholder="Cerca Struttura..."
                        value={storeSearch}
                        onChange={(e) => {
                          setStoreSearch(e.target.value);
                          setShowStoreDropdown(true);
                        }}
                        onFocus={() => setShowStoreDropdown(true)}
                        className="w-full bg-transparent border-b border-[#141414] font-mono font-bold text-xs focus:outline-none py-1"
                      />
                      {showStoreDropdown && (
                        <div className="absolute top-full left-0 w-full max-h-60 overflow-y-auto bg-white border border-[#141414] z-[100] shadow-xl mt-1">
                          <div 
                            className="p-2 hover:bg-gray-100 cursor-pointer text-[10px] uppercase tracking-widest font-bold border-b border-gray-100"
                            onClick={() => {
                              setSelectedStoreId('');
                              setStoreSearch('');
                              setShowStoreDropdown(false);
                            }}
                          >
                            Nessuna Struttura
                          </div>
                          {storesData
                            .filter(s => s.nome.toLowerCase().includes(storeSearch.toLowerCase()))
                            .map(s => (
                              <div 
                                key={s.id}
                                className="p-2 hover:bg-gray-100 cursor-pointer text-[10px] font-mono border-b border-gray-100 last:border-0"
                                onClick={() => {
                                  setSelectedStoreId(s.id);
                                  setStoreSearch(s.nome);
                                  setShowStoreDropdown(false);
                                }}
                              >
                                {s.nome}
                              </div>
                            ))
                          }
                        </div>
                      )}
                    </div>
                  </div>
                  {analyzedPatients.length > 0 && (
                    <div className="flex items-center gap-2">
                      <span className="text-[10px] uppercase tracking-widest opacity-60 font-mono">
                        {analyzedPatients.length} pazienti trovati
                      </span>
                      <button
                        onClick={handleDownloadResults}
                        className="p-2 hover:bg-gray-100 rounded-full transition-colors text-emerald-600"
                        title="Scarica Excel"
                      >
                        <Download size={18} />
                      </button>
                    </div>
                  )}
                  <div className="flex flex-wrap items-center gap-2">
                    <div className="flex items-center border border-[#141414]">
                      <select
                        value={selectedMonths}
                        onChange={(e) => setSelectedMonths(parseInt(e.target.value))}
                        className="bg-transparent font-mono font-bold text-[10px] focus:outline-none cursor-pointer px-2 py-2 border-r border-[#141414]"
                      >
                        {[1, 2, 3, 4, 5, 6].map(m => (
                          <option key={m} value={m}>{m} mesi</option>
                        ))}
                      </select>
                      <button
                        onClick={handleNoContattatiDa}
                        disabled={isAnalyzing || patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0 || (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim())}
                        className={cn(
                          "px-4 py-2 text-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all",
                          (isAnalyzing || patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0 || (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim())) 
                            ? "opacity-30 cursor-not-allowed" 
                            : "hover:bg-gray-100 active:scale-95"
                        )}
                      >
                        No contattati da...
                      </button>
                    </div>
                  <div className="flex items-center gap-2">
                    <div className="flex items-center gap-1">
                      <div className="relative group">
                        <button
                          onClick={handleInteressanti}
                          disabled={isAnalyzing || patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0 || (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim())}
                          className={cn(
                            "px-6 py-2 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold transition-all",
                            (isAnalyzing || patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0 || (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim())) 
                              ? "opacity-30 cursor-not-allowed" 
                              : "hover:bg-[#333] active:scale-95"
                          )}
                        >
                          {isAnalyzing ? "Analisi..." : "Interessanti"}
                        </button>
                        
                        {/* Legend Popover moved to trigger on icon hover specifically for better UX */}
                      </div>
                      
                      <div className="relative group/info p-1">
                        <Info size={18} className="text-[#141414] opacity-50 hover:opacity-100 cursor-help transition-all" />
                        
                        <div className="absolute top-full left-1/2 -translate-x-1/2 mt-3 w-80 bg-white border border-[#141414] shadow-2xl p-4 hidden group-hover/info:block z-[999] pointer-events-none">
                          <div className="flex items-center gap-2 mb-3 border-b border-gray-100 pb-2">
                            <Info size={14} className="text-[#141414]" />
                            <span className="text-[10px] uppercase tracking-widest font-bold">Logica Estrazione "Interessanti"</span>
                          </div>
                          
                          <div className="space-y-4">
                            <div>
                              <p className="text-[9px] uppercase tracking-wider font-bold mb-1.5 opacity-60 text-red-600 border-b border-red-50 w-fit">1. Filtri di Esclusione</p>
                              <ul className="text-[10px] space-y-1 list-disc pl-3 opacity-80 decoration-red-200">
                                <li><span className="font-bold">Status:</span> Deceduti, Clienti, Errori, Normoudenti</li>
                                <li><span className="font-bold">Recentismo App.:</span> Futuri o <span className="text-red-700 font-bold">ultimi 2 mesi</span></li>
                                <li><span className="font-bold">Recentismo Call:</span> Chiamate negli <span className="text-red-700 font-bold">ultimi 3 mesi</span></li>
                                <li><span className="font-bold">Keyword Negative:</span> Note con <span className="italic">"non interessato"</span></li>
                              </ul>
                            </div>

                            <div>
                              <p className="text-[9px] uppercase tracking-wider font-bold mb-1.5 opacity-60 text-[#141414] border-b border-gray-100 w-fit">2. Filtro Area</p>
                              <p className="text-[10px] opacity-80">Solo nominativi appartenenti alla <span className="font-bold">Filiale</span>, <span className="font-bold">Struttura</span> o <span className="font-bold">CAP</span> selezionati nei filtri sopra.</p>
                            </div>
                            
                            <div>
                              <p className="text-[9px] uppercase tracking-wider font-bold mb-1.5 opacity-60 text-emerald-600 border-b border-emerald-50 w-fit">3. Scoring e Inclusione</p>
                              <p className="text-[10px] mb-2 opacity-80 italic">Vengono inclusi solo i nominativi con almeno una parola chiave trovata nelle note (Anagrafica, Call o Appuntamenti):</p>
                              <div className="space-y-2">
                                <div className="bg-emerald-50/50 p-2 border-l-2 border-emerald-500">
                                  <p className="font-bold text-emerald-700 text-[9px]">ALTA PRIORITÀ (+10 pt):</p>
                                  <p className="text-[9px] opacity-70 italic">Keywords: apparecchi, protesi, sentire meglio, prova</p>
                                </div>
                                <div className="bg-blue-50/50 p-2 border-l-2 border-blue-500">
                                  <p className="font-bold text-blue-700 text-[9px]">MEDIA PRIORITÀ (+5 pt):</p>
                                  <p className="text-[9px] opacity-70 italic">Keywords: ipoacusico, visita, controllo, esame, udito</p>
                                </div>
                              </div>
                            </div>
                            
                            <div className="pt-2 border-t border-gray-100 text-[9px] italic opacity-50 leading-tight">
                              * Il sistema scansiona l'intero storico delle note per ogni paziente filtrato.
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    <button
                      onClick={handleMaiVistiOChiamati}
                      disabled={isAnalyzing || patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0 || (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim())}
                      className={cn(
                        "px-6 py-2 border border-[#141414] text-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all",
                        (isAnalyzing || patientsData.length === 0 || callsData.length === 0 || appointmentsData.length === 0 || (!selectedBranch && !selectedStoreId && !selectedCap && !capFilterSearch.trim())) 
                          ? "opacity-30 cursor-not-allowed" 
                          : "hover:bg-gray-100 active:scale-95"
                      )}
                    >
                      Mai visti o chiamati
                    </button>
                  </div>
                    <button
                      onClick={handleResetPatients}
                      className="px-6 py-2 border border-[#141414] text-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all hover:bg-gray-100 active:scale-95"
                    >
                      Reset
                    </button>
                    <button
                      onClick={handleGeocodePatients}
                      disabled={isGeocoding || patientsData.length === 0}
                      className={cn(
                        "px-6 py-2 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                        (isGeocoding || patientsData.length === 0) 
                          ? "opacity-30 cursor-not-allowed" 
                          : "hover:bg-[#333] active:scale-95"
                      )}
                    >
                      {isGeocoding ? <Loader2 className="w-3 h-3 animate-spin" /> : <MapIcon className="w-3 h-3" />}
                      Seleziona da Mappa
                    </button>
                  </div>
                </div>
              </div>
              
              {showMap && (
                <div className="fixed inset-0 z-50 bg-white flex flex-col">
                  <div className="p-4 border-bottom border-[#141414] flex justify-between items-center bg-white">
                    <div>
                      <h2 className="text-xl font-bold uppercase tracking-tighter">Mappa Pazienti e Strutture</h2>
                      <p className="text-[10px] opacity-60 uppercase tracking-widest">
                        {isGeocoding 
                          ? `Geocodifica in corso... ${geocodingProgress.current}/${geocodingProgress.total}` 
                          : `Visualizzati ${geocodedPatients.length} pazienti e ${geocodedStores.length} strutture/filiali.`}
                      </p>
                    </div>
                    <div className="flex items-center gap-4">
                      <button
                        onClick={handleDownloadMap}
                        className="flex items-center gap-2 px-3 py-2 bg-white border border-gray-200 text-blue-600 text-[10px] uppercase tracking-widest font-bold hover:bg-gray-50 transition-all rounded shadow-sm"
                        title="Scarica immagine mappa"
                      >
                        <Download size={14} />
                        Mappa
                      </button>
                      {analyzedPatients.length > 0 && (
                        <button
                          onClick={handleDownloadResults}
                          className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white text-[10px] uppercase tracking-widest font-bold hover:bg-emerald-700 transition-all rounded shadow-sm"
                        >
                          <Download size={14} />
                          Scarica Selezione ({analyzedPatients.length})
                        </button>
                      )}
                      <button 
                        onClick={() => setShowMap(false)}
                        className="p-2 hover:bg-gray-100 rounded-full transition-colors"
                      >
                        <X size={18} />
                      </button>
                    </div>
                  </div>
                  <div className="flex-1 relative" ref={mapRef}>
                    <MapContainer 
                      center={[41.9028, 12.4964]} 
                      zoom={6} 
                      style={{ height: '100%', width: '100%' }}
                    >
                      <TileLayer
                        attribution='&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                        url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
                        crossOrigin=""
                      />
                      <GeomanControl onSelection={handleMapSelection} />
                      {/* Render Stores */}
                      {geocodedStores
                        .filter(s => !["Roma - Sede", "Roma - Villa Chigi", "Roma - San Giovanni"].includes(String(s.nome || '').trim()))
                        .map((s, idx) => (
                        <Marker 
                          key={`store-${idx}`} 
                          position={[s.position.lat, s.position.lng]}
                          icon={
                            (String(s.tipo || '').toLowerCase().includes('corporate') || String(s.tipo || '').toLowerCase().includes('franchising')) ? largeGreenIcon :
                            String(s.tipo || '').toLowerCase().includes('filiale') ? blueIcon : redIcon
                          }
                        >
                          <Popup>
                            <div className="p-1">
                              <div className="text-[8px] uppercase tracking-widest font-bold text-red-600 mb-1">
                                {s.tipo || 'Struttura'}
                              </div>
                              <h4 className="font-bold text-sm">{s.nome}</h4>
                              <p className="text-xs opacity-70">{s.indirizzo}</p>
                              <p className="text-xs opacity-70">{s.cap} {s.citta}</p>
                              {s.note && <p className="text-[10px] mt-1 border-t pt-1 opacity-60 italic">{s.note}</p>}
                            </div>
                          </Popup>
                        </Marker>
                      ))}
                      {/* Render Patients */}
                      {geocodedPatients.map((p, idx) => (
                        <Marker 
                          key={`patient-${idx}`} 
                          position={[p.position.lat, p.position.lng]}
                          icon={defaultIcon}
                        >
                          <Popup>
                            <div className="p-1">
                              <h4 className="font-bold text-sm">{p.nomeCompleto}</h4>
                              <p className="text-xs opacity-70">{p.indirizzo}</p>
                              <p className="text-xs opacity-70">{p.cap} {p.citta}</p>
                              {p.tipo && (
                                <span className="mt-1 inline-block text-[8px] bg-blue-50 text-blue-700 px-1 py-0.5 rounded border border-blue-100 uppercase font-bold">
                                  {p.tipo}
                                </span>
                              )}
                            </div>
                          </Popup>
                        </Marker>
                      ))}
                    </MapContainer>
                  </div>
                </div>
              )}

              <div className="divide-y divide-[#141414]">
                {analyzedPatients.length > 0 ? (
                  analyzedPatients.map((p, idx) => (
                    <div key={idx} className="p-6 hover:bg-gray-50 transition-colors flex flex-col md:flex-row justify-between gap-6">
                      <div className="space-y-3 flex-1">
                        <div className="flex items-center gap-3">
                          <h3 className="text-lg font-bold">{p.nomeCompleto}</h3>
                          <span className="text-[10px] font-mono bg-gray-100 px-2 py-0.5 rounded">ID: {p.id || 'N/D'}</span>
                          {p.tipo && (
                            <span className="text-[10px] font-mono bg-blue-50 text-blue-700 px-2 py-0.5 rounded border border-blue-100 uppercase tracking-widest font-bold">
                              {p.tipo}
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-3 text-[10px] opacity-60 font-serif italic">
                          <span>Città: {p.citta}</span>
                          {p.cap && p.cap !== 'N/D' && (
                            <span className="bg-gray-100 text-[#141414] px-2 py-0.5 rounded font-mono font-bold not-italic">CAP: {p.cap}</span>
                          )}
                        </div>
                        <div className="flex flex-wrap gap-2">
                          {p.reasons.map((r: string, i: number) => (
                            <span key={i} className="text-[9px] uppercase tracking-widest px-2 py-1 bg-emerald-100 text-emerald-800 font-bold">
                              {r}
                            </span>
                          ))}
                        </div>
                        <div className="text-xs opacity-70 italic font-serif">
                          "{p.note || 'Nessuna nota in anagrafica'}"
                        </div>
                      </div>

                      <div className="flex flex-col justify-between items-end gap-4 min-w-[200px]">
                        <div className="text-right space-y-1">
                          <div className="text-[10px] uppercase tracking-widest opacity-40">Contatti</div>
                          <div className="text-sm font-mono font-bold">{p.telefono || 'Nessun numero'}</div>
                        </div>
                        <div className="grid grid-cols-2 gap-4 text-right">
                          <div>
                            <div className="text-[9px] uppercase tracking-widest opacity-40">Ultima Tel.</div>
                            <div className="text-[10px] font-mono">{p.lastCall}</div>
                          </div>
                          <div>
                            <div className="text-[9px] uppercase tracking-widest opacity-40">Ultimo App.</div>
                            <div className="text-[10px] font-mono">{p.lastApp}</div>
                          </div>
                        </div>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="p-12 text-center opacity-40 italic font-serif text-lg">
                    {patientsData.length > 0 && callsData.length > 0 && appointmentsData.length > 0 
                      ? "Premi 'Esamina' per generare la lista prioritaria"
                      : "Carica i file ii.xlsx, tt.xlsx e aa.xlsx per iniziare l'analisi"}
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : activeTab === 'appointmentsManagement' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <Calendar size={32} /> Gestione Appuntamenti
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Analisi degli appuntamenti confermati (Si è presentato) per audioprotesista e {appointmentsAnalysisView === 'canali' ? 'canale marketing' : 'tipo di contatto'}.
              </p>
            </div>
            <div className="flex items-center gap-4">
              <div className="flex items-center bg-gray-100 p-1 border border-[#141414] mt-5">
                <button
                  onClick={() => setAppointmentsAnalysisView('canali')}
                  className={`px-4 py-2 text-[10px] uppercase tracking-widest font-bold transition-all ${appointmentsAnalysisView === 'canali' ? 'bg-[#141414] text-white' : 'text-gray-500 hover:text-black'}`}
                >
                  Canali
                </button>
                <button
                  onClick={() => setAppointmentsAnalysisView('tipi')}
                  className={`px-4 py-2 text-[10px] uppercase tracking-widest font-bold transition-all ${appointmentsAnalysisView === 'tipi' ? 'bg-[#141414] text-white' : 'text-gray-500 hover:text-black'}`}
                >
                  Tipi
                </button>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-[10px] uppercase tracking-widest opacity-50">Anno di riferimento</label>
                <div className="flex items-center border border-[#141414] bg-white">
                  <Calendar size={16} className="ml-3 opacity-40" />
                  <select
                    value={selectedApptYear}
                    onChange={(e) => setSelectedApptYear(parseInt(e.target.value))}
                    className="bg-transparent font-mono font-bold text-xs focus:outline-none cursor-pointer px-4 py-3 min-w-[120px]"
                  >
                    {availableApptYears.map(y => (
                      <option key={y} value={y}>{y}</option>
                    ))}
                  </select>
                </div>
              </div>
              <button
                onClick={downloadAppointmentsExcel}
                disabled={appointmentsAnalysis.length === 0}
                className="bg-[#141414] text-white px-6 py-3 text-[10px] uppercase tracking-widest font-bold hover:bg-black transition-all flex items-center gap-2 h-[41px] mt-5 self-end disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Download size={14} /> Download Excel
              </button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full border-collapse border border-[#141414] text-[10px]">
              <thead>
                <tr className="bg-[#141414] text-white uppercase tracking-widest">
                  <th className="border border-[#141414] p-2 text-left w-64">{appointmentsAnalysisView === 'canali' ? 'Audioprotesista / Canale' : 'Audioprotesista / Tipo'}</th>
                  {MONTH_NAMES.map(m => (
                    <th key={m} className="border border-[#141414] p-2 text-center w-16">{m}</th>
                  ))}
                  <th className="border border-[#141414] p-2 text-center w-20">Totale</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#141414]">
                {appointmentsAnalysis.length > 0 ? (
                  appointmentsAnalysis.map((pro, idx) => (
                    <React.Fragment key={idx}>
                      <tr className="bg-gray-200 font-bold">
                        <td className="border border-[#141414] p-3 text-sm flex items-center gap-2">
                          <Users size={14} /> {pro.name}
                        </td>
                        {pro.months.map((count, mIdx) => (
                          <td 
                            key={mIdx} 
                            className={`border border-[#141414] p-3 text-center text-sm ${count > 0 ? 'cursor-pointer hover:bg-gray-300 hover:text-blue-600 transition-colors' : ''}`}
                            onClick={() => count > 0 && handleAppointmentsDrillDown(pro.name, mIdx)}
                          >
                            {count || '-'}
                          </td>
                        ))}
                        <td 
                          className={`border border-[#141414] p-3 text-center text-sm bg-gray-300 ${pro.months.reduce((a, b) => a + b, 0) > 0 ? 'cursor-pointer hover:bg-gray-400 hover:text-blue-600 transition-colors' : ''}`}
                          onClick={() => pro.months.reduce((a, b) => a + b, 0) > 0 && handleAppointmentsDrillDown(pro.name, -1)}
                        >
                          {pro.months.reduce((a, b) => a + b, 0)}
                        </td>
                      </tr>
                      {Array.from(pro.subGroups.entries()).map(([subGroup, subGroupMonths], cIdx) => (
                        <tr key={`${idx}-${cIdx}`} className="hover:bg-gray-50 border-b border-gray-100 italic">
                          <td className="border border-[#141414] p-2 pl-8 opacity-70">
                            {subGroup}
                          </td>
                          {subGroupMonths.map((count, mIdx) => (
                            <td 
                              key={mIdx} 
                              className={`border border-[#141414] p-2 text-center opacity-70 ${count > 0 ? 'cursor-pointer hover:bg-gray-200 hover:text-blue-600 transition-colors' : ''}`}
                              onClick={() => count > 0 && handleAppointmentsDrillDown(pro.name, mIdx, subGroup)}
                            >
                              {count || '-'}
                            </td>
                          ))}
                          <td 
                            className={`border border-[#141414] p-2 text-center opacity-70 bg-gray-50 ${subGroupMonths.reduce((a, b) => a + b, 0) > 0 ? 'cursor-pointer hover:bg-gray-200 hover:text-blue-600 transition-colors' : ''}`}
                            onClick={() => subGroupMonths.reduce((a, b) => a + b, 0) > 0 && handleAppointmentsDrillDown(pro.name, -1, subGroup)}
                          >
                            {subGroupMonths.reduce((a, b) => a + b, 0)}
                          </td>
                        </tr>
                      ))}
                    </React.Fragment>
                  ))
                ) : (
                  <tr>
                    <td colSpan={14} className="p-12 text-center border border-[#141414] opacity-40 italic font-serif text-lg">
                      {appointmentsData.length > 0 
                        ? "Nessun appuntamento trovato per i criteri selezionati"
                        : "Carica il file aa.xlsx per visualizzare l'analisi degli appuntamenti"}
                    </td>
                  </tr>
                )}
              </tbody>
              {appointmentsAnalysis.length > 0 && (
                <tfoot className="bg-[#141414] text-white font-bold uppercase tracking-widest text-[10px]">
                  <tr>
                    <td className="border border-[#141414] p-3 text-left">TOTALE GENERALE</td>
                    {Array.from({ length: 12 }).map((_, mIdx) => {
                      const monthTotal = appointmentsAnalysis.reduce((sum, pro) => sum + (pro.months[mIdx] || 0), 0);
                      return (
                        <td 
                          key={mIdx} 
                          className={`border border-[#141414] p-3 text-center ${monthTotal > 0 ? 'cursor-pointer hover:bg-gray-800 hover:text-blue-400 transition-colors' : ''}`}
                          onClick={() => monthTotal > 0 && handleAppointmentsDrillDown(null, mIdx)}
                        >
                          {monthTotal || '-'}
                        </td>
                      );
                    })}
                    <td 
                      className="border border-[#141414] p-3 text-center bg-gray-700 cursor-pointer hover:bg-gray-600 hover:text-blue-400 transition-colors"
                      onClick={() => {
                        const total = appointmentsAnalysis.reduce((total, pro) => total + pro.months.reduce((a, b) => a + b, 0), 0);
                        if (total > 0) handleAppointmentsDrillDown(null, -1);
                      }}
                    >
                      {appointmentsAnalysis.reduce((total, pro) => total + pro.months.reduce((a, b) => a + b, 0), 0)}
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>
      ) : activeTab === 'screeningEntrances' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <Users size={32} /> Ingressi Screening
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Analisi degli appuntamenti confermati (Si è presentato) per audioprotesista e filiale (sede).
              </p>
            </div>
            <div className="flex items-center gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-[10px] uppercase tracking-widest opacity-50">Anno di riferimento</label>
                <div className="flex items-center border border-[#141414] bg-white">
                  <Calendar size={16} className="ml-3 opacity-40" />
                  <select
                    value={selectedScreeningYear}
                    onChange={(e) => setSelectedScreeningYear(parseInt(e.target.value))}
                    className="bg-transparent font-mono font-bold text-xs focus:outline-none cursor-pointer px-4 py-3 min-w-[120px]"
                  >
                    {availableApptYears.map(y => (
                      <option key={y} value={y}>{y}</option>
                    ))}
                  </select>
                </div>
              </div>
              <button
                onClick={downloadScreeningEntrancesExcel}
                disabled={screeningEntrancesAnalysis.length === 0}
                className="bg-[#141414] text-white px-6 py-3 text-[10px] uppercase tracking-widest font-bold hover:bg-black transition-all flex items-center gap-2 h-[41px] mt-5 self-end disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Download size={14} /> Download Excel
              </button>
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full border-collapse border border-[#141414] text-[10px]">
              <thead>
                <tr className="bg-[#141414] text-white uppercase tracking-widest">
                  <th className="border border-[#141414] p-2 text-left w-64">Audioprotesista / Store</th>
                  {MONTH_NAMES.map(m => (
                    <th key={m} className="border border-[#141414] p-2 text-center w-16">{m}</th>
                  ))}
                  <th className="border border-[#141414] p-2 text-center w-20">Totale</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#141414]">
                {screeningEntrancesAnalysis.length > 0 ? (
                  screeningEntrancesAnalysis.map((pro, idx) => (
                    <React.Fragment key={idx}>
                      <tr className="bg-gray-200 font-bold">
                        <td className="border border-[#141414] p-3 text-sm flex items-center gap-2">
                          <Users size={14} /> {pro.name}
                        </td>
                        {pro.months.map((count, mIdx) => (
                          <td 
                            key={mIdx} 
                            className={`border border-[#141414] p-3 text-center text-sm ${count > 0 ? 'cursor-pointer hover:bg-gray-300 hover:text-blue-600 transition-colors' : ''}`}
                            onClick={() => count > 0 && handleScreeningDrillDown(pro.name, mIdx)}
                          >
                            {count || '-'}
                          </td>
                        ))}
                        <td 
                          className={`border border-[#141414] p-3 text-center text-sm bg-gray-300 ${pro.months.reduce((a, b) => a + b, 0) > 0 ? 'cursor-pointer hover:bg-gray-400 hover:text-blue-600 transition-colors' : ''}`}
                          onClick={() => pro.months.reduce((a, b) => a + b, 0) > 0 && handleScreeningDrillDown(pro.name, -1)}
                        >
                          {pro.months.reduce((a, b) => a + b, 0)}
                        </td>
                      </tr>
                      {Array.from(pro.stores.entries()).map(([store, storeMonths], sIdx) => (
                        <tr key={`${idx}-${sIdx}`} className="hover:bg-gray-50 border-b border-gray-100 italic">
                          <td className="border border-[#141414] p-2 pl-8 opacity-70">
                            {store}
                          </td>
                          {storeMonths.map((count, mIdx) => (
                            <td 
                              key={mIdx} 
                              className={`border border-[#141414] p-2 text-center opacity-70 ${count > 0 ? 'cursor-pointer hover:bg-gray-200 hover:text-blue-600 transition-colors' : ''}`}
                              onClick={() => count > 0 && handleScreeningDrillDown(pro.name, mIdx, store)}
                            >
                              {count || '-'}
                            </td>
                          ))}
                          <td 
                            className={`border border-[#141414] p-2 text-center opacity-70 bg-gray-50 ${storeMonths.reduce((a, b) => a + b, 0) > 0 ? 'cursor-pointer hover:bg-gray-200 hover:text-blue-600 transition-colors' : ''}`}
                            onClick={() => storeMonths.reduce((a, b) => a + b, 0) > 0 && handleScreeningDrillDown(pro.name, -1, store)}
                          >
                            {storeMonths.reduce((a, b) => a + b, 0)}
                          </td>
                        </tr>
                      ))}
                    </React.Fragment>
                  ))
                ) : (
                  <tr>
                    <td colSpan={14} className="p-12 text-center border border-[#141414] opacity-40 italic font-serif text-lg">
                      {appointmentsData.length > 0 
                        ? "Nessun appuntamento trovato per i criteri selezionati"
                        : "Carica il file aa.xlsx per visualizzare l'analisi degli ingressi screening"}
                    </td>
                  </tr>
                )}
              </tbody>
              {screeningEntrancesAnalysis.length > 0 && (
                <tfoot className="bg-[#141414] text-white font-bold uppercase tracking-widest text-[10px]">
                  <tr>
                    <td className="border border-[#141414] p-3 text-left">TOTALE GENERALE</td>
                    {Array.from({ length: 12 }).map((_, mIdx) => {
                      const monthTotal = screeningEntrancesAnalysis.reduce((sum, pro) => sum + (pro.months[mIdx] || 0), 0);
                      return (
                        <td 
                          key={mIdx} 
                          className={`border border-[#141414] p-3 text-center ${monthTotal > 0 ? 'cursor-pointer hover:bg-gray-800 hover:text-blue-400 transition-colors' : ''}`}
                          onClick={() => monthTotal > 0 && handleScreeningDrillDown(null, mIdx)}
                        >
                          {monthTotal || '-'}
                        </td>
                      );
                    })}
                    <td 
                      className="border border-[#141414] p-3 text-center bg-gray-700 cursor-pointer hover:bg-gray-600 hover:text-blue-400 transition-colors"
                      onClick={() => {
                        const total = screeningEntrancesAnalysis.reduce((total, pro) => total + pro.months.reduce((a, b) => a + b, 0), 0);
                        if (total > 0) handleScreeningDrillDown(null, -1);
                      }}
                    >
                      {screeningEntrancesAnalysis.reduce((total, pro) => total + pro.months.reduce((a, b) => a + b, 0), 0)}
                    </td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>
      ) : activeTab === 'recapiti' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <Building2 size={32} /> Pazienti-Recapiti
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Seleziona uno o più recapiti esterni per scaricare le liste dei pazienti non contattati.
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-4">
              <div className="flex items-center border border-[#141414] bg-white">
                <Calendar size={16} className="ml-3 opacity-40" />
                <select
                  value={availabilityMonth}
                  onChange={(e) => setAvailabilityMonth(e.target.value as 'current' | 'next')}
                  className="bg-transparent font-mono font-bold text-[10px] focus:outline-none cursor-pointer px-4 py-3 border-r border-[#141414]"
                >
                  <option value="current">Questo mese</option>
                  <option value="next">Mese successivo</option>
                </select>
                <button
                  onClick={handleSelectRecapitiByAvailability}
                  className="px-4 py-3 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold hover:bg-[#333] transition-all"
                >
                  Seleziona i recapiti
                </button>
                <button
                  onClick={() => {
                    setSelectedRecapitiIds([]);
                    setRecapitiDates({});
                  }}
                  className="px-4 py-3 bg-white text-[#141414] border-l border-[#141414] text-[10px] uppercase tracking-widest font-bold hover:bg-gray-100 transition-all"
                >
                  Reset
                </button>
              </div>
              <div className="flex items-center border border-[#141414] bg-white">
                <Search size={16} className="ml-3 opacity-40" />
                <input 
                  type="text"
                  placeholder="Cerca recapito..."
                  value={capSearch}
                  onChange={(e) => {
                    setCapSearch(e.target.value);
                    setCapPage(1);
                  }}
                  className="bg-transparent font-mono font-bold text-xs focus:outline-none px-3 py-3 w-48"
                />
              </div>
              <div className="flex items-center border border-[#141414]">
                <select
                  value={selectedMonths}
                  onChange={(e) => setSelectedMonths(parseInt(e.target.value))}
                  className="bg-transparent font-mono font-bold text-[10px] focus:outline-none cursor-pointer px-4 py-3 border-r border-[#141414]"
                >
                  {[1, 2, 3, 4, 5, 6].map(m => (
                    <option key={m} value={m}>{m} mesi</option>
                  ))}
                </select>
                <button
                  onClick={handleDownloadRecapitiLists}
                  disabled={isAnalyzing || selectedRecapitiIds.length === 0 || patientsData.length === 0}
                  className={cn(
                    "px-6 py-3 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                    (isAnalyzing || selectedRecapitiIds.length === 0 || patientsData.length === 0) ? "opacity-30 cursor-not-allowed" : "hover:bg-[#333] active:scale-95"
                  )}
                >
                  {isAnalyzing ? <Loader2 size={14} className="animate-spin" /> : <Download size={14} />}
                  Scarica Liste ({selectedRecapitiIds.length})
                </button>
              </div>
            </div>
          </div>

          <div className="overflow-x-auto mb-8">
            <table className="w-full text-left border-collapse border border-[#141414]">
              <thead>
                <tr className="bg-[#141414] text-white">
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20 w-12 text-center">
                    <input 
                      type="checkbox" 
                      onChange={(e) => {
                        if (e.target.checked) {
                          setSelectedRecapitiIds(filteredStores.map(s => s.id));
                        } else {
                          setSelectedRecapitiIds([]);
                        }
                      }}
                      checked={filteredStores.length > 0 && selectedRecapitiIds.length === filteredStores.length}
                    />
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Nome Recapito</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Città</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Data Screening</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold">CAP Associati</th>
                </tr>
              </thead>
              <tbody className="font-mono text-xs">
                {filteredStores.length > 0 ? (
                  filteredStores
                    .map((s, idx) => {
                      const associatedCaps = storeCaps
                        .filter(sc => sc.storeId === s.id)
                        .map(sc => sc.cap);
                      
                      return (
                        <tr key={s.id || idx} className={cn(
                          "border-b border-[#141414] hover:bg-gray-50 transition-colors",
                          selectedRecapitiIds.includes(s.id) ? "bg-emerald-50/30" : ""
                        )}>
                          <td className="p-4 border-r border-[#141414] text-center">
                            <input 
                              type="checkbox" 
                              checked={selectedRecapitiIds.includes(s.id)}
                              onChange={() => {
                                setSelectedRecapitiIds(prev => 
                                  prev.includes(s.id) ? prev.filter(id => id !== s.id) : [...prev, s.id]
                                );
                              }}
                            />
                          </td>
                          <td className="p-4 border-r border-[#141414] font-bold">{s.nome}</td>
                          <td className="p-4 border-r border-[#141414]">{s.citta}</td>
                          <td className="p-4 border-r border-[#141414] font-bold text-emerald-700">
                            {recapitiDates[s.id] || '-'}
                          </td>
                          <td className="p-4">
                            <div className="flex flex-wrap gap-1">
                              {s.cap && (
                                <span className="bg-gray-100 text-gray-800 px-2 py-0.5 rounded text-[9px] font-bold border border-gray-200">
                                  {s.cap}
                                </span>
                              )}
                              {associatedCaps.map((cap, cIdx) => (
                                <span key={cIdx} className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded text-[9px] font-bold border border-emerald-200">
                                  {cap}
                                </span>
                              ))}
                            </div>
                          </td>
                        </tr>
                      );
                    })
                ) : (
                  <tr>
                    <td colSpan={5} className="p-12 text-center opacity-40 italic font-serif text-lg">
                      Nessun recapito trovato. Carica il file delle strutture nella sezione Gestione Pazienti.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      ) : activeTab === 'verifiche' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <CheckCircle2 size={32} /> Verifiche Anagrafiche
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Strumenti di controllo per la qualità dei dati e la coerenza delle anagrafiche.
              </p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button
                onClick={findDuplicates}
                className={cn(
                  "px-4 py-3 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                  activeVerification === 'doppioni' ? "bg-[#141414] text-white" : "bg-white text-[#141414] hover:bg-gray-50"
                )}
              >
                <Copy size={14} /> Doppioni
              </button>
              <button
                onClick={findNoCap}
                className={cn(
                  "px-4 py-3 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                  activeVerification === 'nocap' ? "bg-[#141414] text-white" : "bg-white text-[#141414] hover:bg-gray-50"
                )}
              >
                <MapPin size={14} /> No Cap
              </button>
              <button
                onClick={findNoPhone}
                className={cn(
                  "px-4 py-3 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                  activeVerification === 'nophone' ? "bg-[#141414] text-white" : "bg-white text-[#141414] hover:bg-gray-50"
                )}
              >
                <PhoneOff size={14} /> No Telefono
              </button>
              <button
                onClick={findClientsNoProforma}
                className={cn(
                  "px-4 py-3 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                  activeVerification === 'noproforma' ? "bg-[#141414] text-white" : "bg-white text-[#141414] hover:bg-gray-50"
                )}
              >
                <FileX size={14} /> Clienti no proforma
              </button>
              <button
                onClick={findContattiChiamati}
                className={cn(
                  "px-4 py-3 border border-[#141414] text-[10px] uppercase tracking-widest font-bold transition-all flex items-center gap-2",
                  activeVerification === 'contattichiamati' ? "bg-[#141414] text-white" : "bg-white text-[#141414] hover:bg-gray-50"
                )}
              >
                <UserCheck size={14} /> Contatti Chiamati
              </button>
            </div>
          </div>

          <div className="space-y-4">
            {verificationResults.length > 0 ? (
              <div className="border border-[#141414]">
                <div className="bg-gray-100 p-4 border-b border-[#141414] flex justify-between items-center">
                  <span className="text-[10px] uppercase tracking-widest font-bold opacity-60">
                    Risultati: {verificationResults.length} pazienti trovati
                  </span>
                  <button
                    onClick={() => generateAndDownloadExcel(verificationResults, `Verifica_${activeVerification}_${new Date().toISOString().split('T')[0]}.xlsx`)}
                    className="flex items-center gap-2 text-[10px] uppercase tracking-widest font-bold hover:underline"
                  >
                    <Download size={14} /> Scarica Excel
                  </button>
                </div>
                <div className="divide-y divide-gray-200 max-h-[600px] overflow-y-auto">
                  {verificationResults.map((p, idx) => (
                    <div key={idx} className="p-4 hover:bg-gray-50 transition-colors flex justify-between items-center">
                      <div>
                        <div className="font-bold">{p.nomeCompleto}</div>
                        <div className="text-[10px] opacity-60 font-mono">
                          {p.indirizzo} - {p.citta} {p.provincia && `(${p.provincia})`} {p.cap && `[${p.cap}]`}
                        </div>
                        <div className="text-[10px] mt-1">
                          <span className="bg-gray-100 px-2 py-0.5 rounded mr-2">Tipo: {p.tipo}</span>
                          <span className="bg-gray-100 px-2 py-0.5 rounded mr-2">Tel: {p.telefono || 'N/D'}</span>
                          {p.classification && (
                            <div className="mt-2 flex flex-col gap-1">
                              <span className={cn(
                                "px-2 py-0.5 rounded font-bold uppercase tracking-tighter w-fit text-[10px]",
                                p.classification === 'Normoudente' ? "bg-green-100 text-green-700" :
                                p.classification === 'PC' ? "bg-blue-100 text-blue-700" :
                                "bg-orange-100 text-orange-700"
                              )}>
                                Valutazione: {p.classification}
                              </span>
                              {p.classificationReason && (
                                <span className="text-[10px] italic opacity-70">
                                  Motivazione: {p.classificationReason}
                                </span>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                      <div className="text-right">
                        <div className="text-[10px] uppercase tracking-widest opacity-40">Store</div>
                        <div className="text-xs font-bold">{p.store}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : activeVerification ? (
              <div className="p-20 text-center border border-dashed border-gray-300 rounded-lg">
                <div className="text-4xl mb-4">✨</div>
                <div className="font-serif italic text-xl opacity-60">Nessuna anomalia riscontrata per questa verifica.</div>
              </div>
            ) : (
              <div className="p-20 text-center border border-dashed border-gray-300 rounded-lg">
                <div className="font-serif italic text-xl opacity-40 text-lg">Seleziona una verifica per iniziare l'analisi dei dati.</div>
              </div>
            )}
          </div>
        </div>
      ) : activeTab === 'audioAnalysis' ? (
        <div className="flex flex-col gap-6">
          <div className="bg-white border border-[#141414] p-6">
            <div className="flex flex-wrap gap-6 items-end mb-8">
              <div className="flex-1 min-w-[300px]">
                <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-2">Seleziona Audioprotesista</label>
                <select 
                  value={selectedAudioPro}
                  onChange={(e) => setSelectedAudioPro(e.target.value)}
                  className="w-full bg-white border border-[#141414] p-3 font-mono text-sm focus:outline-none focus:ring-1 focus:ring-[#141414]"
                >
                  <option value="">-- Seleziona --</option>
                  {audioProList.map(pro => (
                    <option key={pro} value={pro}>{pro}</option>
                  ))}
                </select>
              </div>
              <div className="w-[120px]">
                <label className="block text-[10px] uppercase tracking-widest opacity-50 mb-2">Anno</label>
                <select 
                  value={selectedAudioYear}
                  onChange={(e) => setSelectedAudioYear(Number(e.target.value))}
                  className="w-full bg-white border border-[#141414] p-3 font-mono text-sm focus:outline-none focus:ring-1 focus:ring-[#141414]"
                >
                  {(() => {
                    const currentYear = new Date().getFullYear();
                    const years = [];
                    for (let y = 2024; y <= currentYear; y++) {
                      years.push(y);
                    }
                    return years.reverse().map(y => (
                      <option key={y} value={y}>{y}</option>
                    ));
                  })()}
                </select>
              </div>
              {audioPerformance && (
                <button 
                  onClick={downloadAudioPerformanceExcel}
                  className="bg-emerald-600 text-white px-5 py-3 font-mono text-[10px] uppercase tracking-widest flex items-center gap-2 hover:bg-emerald-700 transition-all cursor-pointer h-[45px] shadow-sm active:translate-y-px"
                >
                  <Download size={14} /> Scarica Analisi Excel
                </button>
              )}
            </div>

            {!selectedAudioPro ? (
              <div className="py-20 text-center border-2 border-dashed border-gray-100 rounded-lg">
                <Stethoscope size={40} className="mx-auto text-gray-200 mb-4" />
                <p className="font-serif italic text-xl opacity-40">Seleziona un audioprotesista per visualizzare l'analisi</p>
                <p className="text-[10px] uppercase tracking-widest opacity-30 mt-2">Dati estratti dai file dd.xlsx, tt.xlsx e aa.xlsx</p>
              </div>
            ) : !audioPerformance ? (
              <div className="py-20 text-center border-2 border-dashed border-gray-100 rounded-lg">
                <Loader2 size={40} className="mx-auto text-gray-200 mb-4 animate-spin" />
                <p className="font-serif italic text-xl opacity-40">Caricamento analisi in corso...</p>
              </div>
            ) : (
              <div className="space-y-12">
                {/* 1. Tabella Eventi (dd.xlsx) */}
                <section>
                  <div className="flex items-center gap-3 mb-4">
                    <div className="w-8 h-8 bg-black text-white flex items-center justify-center font-mono text-sm">01</div>
                    <h3 className="font-serif italic text-2xl">Tabella Eventi (dd.xlsx)</h3>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse border border-[#141414]">
                      <thead>
                        <tr className="bg-gray-50 border-b border-[#141414] text-[10px] uppercase tracking-widest">
                          <th className="p-3 border-r border-[#141414]">Store</th>
                          {MONTHS.map(m => (
                            <th key={m} className="p-3 text-center border-r border-[#141414]">{m.slice(0, 3)}</th>
                          ))}
                          <th className="p-3 text-center bg-gray-100">TOT</th>
                        </tr>
                      </thead>
                      <tbody className="font-mono text-xs">
                        {audioPerformance.events.length === 0 ? (
                          <tr>
                            <td colSpan={14} className="p-8 text-center italic opacity-40">Nessun evento trovato per questo anno</td>
                          </tr>
                        ) : (
                          <>
                            {audioPerformance.events.map(([store, months]) => {
                              const total = months.reduce((a, b) => a + b, 0);
                              return (
                                <tr key={store} className="border-b border-[#141414] hover:bg-gray-50 group">
                                  <td 
                                    onClick={() => handleAudioDrillDown('events', -1, store)}
                                    className="p-3 border-r border-[#141414] font-bold cursor-pointer transition-all group-hover:text-emerald-700"
                                  >
                                    {store}
                                  </td>
                                  {months.map((count, idx) => (
                                    <td key={idx} 
                                      onClick={() => count > 0 && handleAudioDrillDown('events', idx, store)}
                                      className={cn(
                                        "p-3 text-center border-r border-[#141414] transition-all",
                                        count > 0 ? "font-bold text-emerald-600 bg-emerald-50/20 cursor-pointer hover:bg-emerald-100/40" : "opacity-30"
                                      )}
                                    >
                                      {count || '-'}
                                    </td>
                                  ))}
                                  <td 
                                    onClick={() => total > 0 && handleAudioDrillDown('events', -1, store)}
                                    className={cn(
                                      "p-3 text-center bg-gray-50 font-bold transition-all",
                                      total > 0 ? "cursor-pointer hover:bg-gray-200" : ""
                                    )}
                                  >
                                    {total}
                                  </td>
                                </tr>
                              );
                            })}
                            <tr className="bg-gray-100 font-bold border-t-2 border-[#141414]">
                              <td className="p-3 border-r border-[#141414] uppercase tracking-widest text-[10px]">Totale Eventi</td>
                              {new Array(12).fill(0).map((_, i) => {
                                const colTotal = audioPerformance.events.reduce((acc, curr) => acc + curr[1][i], 0);
                                return (
                                  <td key={i} className="p-3 text-center border-r border-[#141414] text-emerald-800">
                                    {colTotal || '-'}
                                  </td>
                                );
                              })}
                              <td className="p-3 text-center bg-emerald-900 text-white font-bold">
                                {audioPerformance.events.reduce((acc, curr) => acc + curr[1].reduce((a, b) => a + b, 0), 0)}
                              </td>
                            </tr>
                          </>
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>

                {/* 2. Tabella Telefonate (tt.xlsx) */}
                <section>
                  <div className="flex items-center gap-3 mb-4">
                    <div className="w-8 h-8 bg-black text-white flex items-center justify-center font-mono text-sm">02</div>
                    <h3 className="font-serif italic text-2xl">Tabella Telefonate (tt.xlsx)</h3>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse border border-[#141414]">
                      <thead>
                        <tr className="bg-gray-50 border-b border-[#141414] text-[10px] uppercase tracking-widest">
                          <th className="p-3 border-r border-[#141414]">Esito chiamata</th>
                          {MONTHS.map(m => (
                            <th key={m} className="p-3 text-center border-r border-[#141414]">{m.slice(0, 3)}</th>
                          ))}
                          <th className="p-3 text-center bg-gray-100">TOT</th>
                        </tr>
                      </thead>
                      <tbody className="font-mono text-xs">
                        {audioPerformance.calls.length === 0 ? (
                          <tr>
                            <td colSpan={14} className="p-8 text-center italic opacity-40">Nessuna telefonata trovata per questo anno</td>
                          </tr>
                        ) : (
                          <>
                            {audioPerformance.calls.map(([outcome, months]) => {
                              const total = months.reduce((a, b) => a + b, 0);
                              return (
                                <tr key={outcome} className="border-b border-[#141414] hover:bg-gray-50 group">
                                  <td 
                                    onClick={() => handleAudioDrillDown('calls', -1, outcome)}
                                    className="p-3 border-r border-[#141414] font-bold cursor-pointer transition-all group-hover:text-blue-700"
                                  >
                                    {outcome}
                                  </td>
                                  {months.map((count, idx) => (
                                    <td key={idx} 
                                      onClick={() => count > 0 && handleAudioDrillDown('calls', idx, outcome)}
                                      className={cn(
                                        "p-3 text-center border-r border-[#141414] transition-all",
                                        count > 0 ? "font-bold text-blue-600 bg-blue-50/20 cursor-pointer hover:bg-blue-100/40" : "opacity-30"
                                      )}
                                    >
                                      {count || '-'}
                                    </td>
                                  ))}
                                  <td 
                                    onClick={() => total > 0 && handleAudioDrillDown('calls', -1, outcome)}
                                    className={cn(
                                      "p-3 text-center bg-gray-50 font-bold transition-all",
                                      total > 0 ? "cursor-pointer hover:bg-gray-200" : ""
                                    )}
                                  >
                                    {total}
                                  </td>
                                </tr>
                              );
                            })}
                            <tr className="bg-gray-100 font-bold border-t-2 border-[#141414]">
                              <td className="p-3 border-r border-[#141414] uppercase tracking-widest text-[10px]">Totale Telefonate</td>
                              {new Array(12).fill(0).map((_, i) => {
                                const colTotal = audioPerformance.calls.reduce((acc, curr) => acc + curr[1][i], 0);
                                return (
                                  <td key={i} className="p-3 text-center border-r border-[#141414] text-blue-800">
                                    {colTotal || '-'}
                                  </td>
                                );
                              })}
                              <td className="p-3 text-center bg-blue-900 text-white font-bold">
                                {audioPerformance.calls.reduce((acc, curr) => acc + curr[1].reduce((a, b) => a + b, 0), 0)}
                              </td>
                            </tr>
                          </>
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>

                {/* 3. Tabella Appuntamenti (aa.xlsx) */}
                <section>
                  <div className="flex items-center gap-3 mb-4">
                    <div className="w-8 h-8 bg-black text-white flex items-center justify-center font-mono text-sm">03</div>
                    <h3 className="font-serif italic text-2xl">Tabella Appuntamenti (aa.xlsx)</h3>
                    <div className="bg-emerald-100 text-emerald-800 px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider">
                      Esito: Si è presentato
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse border border-[#141414]">
                      <thead>
                        <tr className="bg-gray-50 border-b border-[#141414] text-[10px] uppercase tracking-widest">
                          <th className="p-3 border-r border-[#141414]">Tipo Contatto</th>
                          {MONTHS.map(m => (
                            <th key={m} className="p-3 text-center border-r border-[#141414]">{m.slice(0, 3)}</th>
                          ))}
                          <th className="p-3 text-center bg-gray-100">TOT</th>
                        </tr>
                      </thead>
                      <tbody className="font-mono text-xs">
                        {audioPerformance.appointments.length === 0 ? (
                          <tr>
                            <td colSpan={14} className="p-8 text-center italic opacity-40">Nessun appuntamento "Si è presentato" trovato per questo anno</td>
                          </tr>
                        ) : (
                          <>
                            {audioPerformance.appointments.map(([type, months]) => {
                              const total = months.reduce((a, b) => a + b, 0);
                              return (
                                <tr key={type} className="border-b border-[#141414] hover:bg-gray-50 group">
                                  <td 
                                    onClick={() => handleAudioDrillDown('appointments', -1, type)}
                                    className="p-3 border-r border-[#141414] font-bold cursor-pointer transition-all group-hover:text-violet-700"
                                  >
                                    {type}
                                  </td>
                                  {months.map((count, idx) => (
                                    <td key={idx} 
                                      onClick={() => count > 0 && handleAudioDrillDown('appointments', idx, type)}
                                      className={cn(
                                        "p-3 text-center border-r border-[#141414] transition-all",
                                        count > 0 ? "font-bold text-violet-600 bg-violet-50/20 cursor-pointer hover:bg-violet-100/40" : "opacity-30"
                                      )}
                                    >
                                      {count || '-'}
                                    </td>
                                  ))}
                                  <td 
                                    onClick={() => total > 0 && handleAudioDrillDown('appointments', -1, type)}
                                    className={cn(
                                      "p-3 text-center bg-gray-50 font-bold transition-all",
                                      total > 0 ? "cursor-pointer hover:bg-gray-200" : ""
                                    )}
                                  >
                                    {total}
                                  </td>
                                </tr>
                              );
                            })}
                            <tr className="bg-gray-100 font-bold border-t-2 border-[#141414]">
                              <td className="p-3 border-r border-[#141414] uppercase tracking-widest text-[10px]">Totale Appuntamenti</td>
                              {new Array(12).fill(0).map((_, i) => {
                                const colTotal = audioPerformance.appointments.reduce((acc, curr) => acc + curr[1][i], 0);
                                return (
                                  <td key={i} className="p-3 text-center border-r border-[#141414] text-violet-800">
                                    {colTotal || '-'}
                                  </td>
                                );
                              })}
                              <td className="p-3 text-center bg-violet-900 text-white font-bold">
                                {audioPerformance.appointments.reduce((acc, curr) => acc + curr[1].reduce((a, b) => a + b, 0), 0)}
                              </td>
                            </tr>
                          </>
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>

                {/* 4. Tabella Prove in corso (pp.xlsx) */}
                <section>
                  <div className="flex items-center gap-3 mb-4">
                    <div className="w-8 h-8 bg-black text-white flex items-center justify-center font-mono text-sm">04</div>
                    <h3 className="font-serif italic text-2xl">Prove in corso (pp.xlsx)</h3>
                    <div className="bg-blue-100 text-blue-800 px-3 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider">
                      Stato: Prova Aperta
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse border border-[#141414]">
                      <thead>
                        <tr className="bg-gray-50 border-b border-[#141414] text-[10px] uppercase tracking-widest">
                          <th className="p-3 border-r border-[#141414]">Cliente finale</th>
                          <th className="p-3 text-right">Valore Totale</th>
                        </tr>
                      </thead>
                      <tbody className="font-mono text-xs">
                        {audioPerformance.openTrials.length === 0 ? (
                          <tr>
                            <td colSpan={2} className="p-8 text-center italic opacity-40">Nessuna prova aperta trovata per questo anno</td>
                          </tr>
                        ) : (
                          audioPerformance.openTrials.map((trial, i) => (
                            <tr key={i} 
                              onClick={() => handleAudioDrillDown('trials', -1)}
                              className="border-b border-[#141414] hover:bg-gray-50 cursor-pointer group"
                            >
                              <td className="p-3 border-r border-[#141414] font-bold group-hover:text-blue-600">{trial.cliente}</td>
                              <td className="p-3 text-right font-bold text-sm text-blue-600">
                                € {(trial.valore || 0).toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                      {audioPerformance.openTrials.length > 0 && (
                        <tfoot className="bg-gray-50 font-bold border-t border-[#141414]">
                          <tr>
                            <td className="p-3 text-right uppercase tracking-widest text-[10px]">Totale Prove Aperte</td>
                            <td className="p-3 text-right text-sm text-blue-800">
                              € {audioPerformance.openTrials.reduce((sum, t) => sum + (t.valore || 0), 0).toLocaleString('it-IT', { useGrouping: true, minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                            </td>
                          </tr>
                        </tfoot>
                      )}
                    </table>
                  </div>
                </section>

                {/* 5. Tabella Vendite vs Storni (ff.xlsx / rr.xlsx) */}
                <section>
                  <div className="flex items-center gap-3 mb-4">
                    <div className="w-8 h-8 bg-black text-white flex items-center justify-center font-mono text-sm">05</div>
                    <h3 className="font-serif italic text-2xl">Vendite e Storni (ff.xlsx / rr.xlsx)</h3>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse border border-[#141414]">
                      <thead>
                        <tr className="bg-gray-50 border-b border-[#141414] text-[10px] uppercase tracking-widest">
                          <th className="p-3 border-r border-[#141414]">Tipo</th>
                          {MONTHS.map(m => (
                            <th key={m} className="p-3 text-center border-r border-[#141414]">{m.slice(0, 3)}</th>
                          ))}
                          <th className="p-3 text-center bg-gray-100">TOTALE</th>
                        </tr>
                      </thead>
                      <tbody className="font-mono text-xs">
                        <tr className="border-b border-[#141414] hover:bg-gray-50">
                          <td className="p-3 border-r border-[#141414] font-bold">Vendite (ff.xlsx)</td>
                          {audioPerformance!.finance.sales.map((v: number, i: number) => (
                            <td key={i} 
                              onClick={() => v > 0 && handleAudioDrillDown('sales', i)}
                              className={cn(
                                "p-3 text-center border-r border-[#141414] text-emerald-600 font-bold transition-all",
                                v > 0 ? "cursor-pointer hover:bg-emerald-50" : ""
                              )}
                            >
                              {v > 0 ? `€ ${v.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}` : '-'}
                            </td>
                          ))}
                          <td 
                            onClick={() => audioPerformance!.finance.sales.reduce((a, b) => a + b, 0) > 0 && handleAudioDrillDown('sales', -1)}
                            className={cn(
                              "p-3 text-center bg-emerald-50 text-emerald-700 font-bold transition-all",
                              audioPerformance!.finance.sales.reduce((a, b) => a + b, 0) > 0 ? "cursor-pointer hover:bg-emerald-100" : ""
                            )}
                          >
                            € {audioPerformance!.finance.sales.reduce((a, b) => a + b, 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                          </td>
                        </tr>
                        <tr className="border-b border-[#141414] hover:bg-gray-50">
                          <td className="p-3 border-r border-[#141414] font-bold">Storni (rr.xlsx)</td>
                          {audioPerformance!.finance.credits.map((v: number, i: number) => (
                            <td key={i} 
                              onClick={() => v > 0 && handleAudioDrillDown('credits', i)}
                              className={cn(
                                "p-3 text-center border-r border-[#141414] text-red-500 font-bold transition-all",
                                v > 0 ? "cursor-pointer hover:bg-red-50" : ""
                              )}
                            >
                              {v > 0 ? `- € ${v.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}` : '-'}
                            </td>
                          ))}
                          <td 
                            onClick={() => audioPerformance!.finance.credits.reduce((a, b) => a + b, 0) > 0 && handleAudioDrillDown('credits', -1)}
                            className={cn(
                              "p-3 text-center bg-red-50 text-red-600 font-bold transition-all",
                              audioPerformance!.finance.credits.reduce((a, b) => a + b, 0) > 0 ? "cursor-pointer hover:bg-red-100" : ""
                            )}
                          >
                            - € {audioPerformance!.finance.credits.reduce((a, b) => a + b, 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                          </td>
                        </tr>
                        <tr className="bg-gray-100 font-bold">
                          <td 
                            onClick={() => (audioPerformance!.finance.sales.reduce((a, b) => a + b, 0) > 0 || audioPerformance!.finance.credits.reduce((a, b) => a + b, 0) > 0) && handleAudioDrillDown('balance', -1)}
                            className="p-3 border-r border-[#141414] uppercase tracking-widest text-[10px] cursor-pointer hover:bg-gray-200 transition-all font-bold"
                          >
                            Saldo Netto
                          </td>
                          {audioPerformance!.finance.balance.map((v: number, i: number) => (
                            <td key={i} 
                              onClick={() => (audioPerformance!.finance.sales[i] > 0 || audioPerformance!.finance.credits[i] > 0) && handleAudioDrillDown('balance', i)}
                              className={cn(
                                "p-3 text-center border-r border-[#141414] transition-all",
                                v >= 0 ? "text-emerald-700" : "text-red-700",
                                (audioPerformance!.finance.sales[i] > 0 || audioPerformance!.finance.credits[i] > 0) ? "cursor-pointer hover:bg-gray-200" : ""
                              )}
                            >
                              € {v.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                            </td>
                          ))}
                          <td 
                            onClick={() => (audioPerformance!.finance.sales.reduce((a, b) => a + b, 0) > 0 || audioPerformance!.finance.credits.reduce((a, b) => a + b, 0) > 0) && handleAudioDrillDown('balance', -1)}
                            className={cn(
                              "p-3 text-center font-bold text-sm transition-all",
                              audioPerformance!.finance.balance.reduce((a, b) => a + b, 0) >= 0 ? "text-emerald-800 bg-emerald-100" : "text-red-800 bg-red-100",
                              (audioPerformance!.finance.sales.reduce((a, b) => a + b, 0) > 0 || audioPerformance!.finance.credits.reduce((a, b) => a + b, 0) > 0) ? "cursor-pointer hover:opacity-80" : ""
                            )}
                          >
                            € {audioPerformance!.finance.balance.reduce((a, b) => a + b, 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </section>
              </div>
            )}
          </div>
        </div>
      ) : activeTab === 'channels' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <TrendingUp size={32} /> Analisi Canali
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Analisi dettagliata delle performance per canale primario: anagrafiche, telefonate, appuntamenti e fatturato.
              </p>
            </div>
            <div className="flex flex-col md:flex-row items-end gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-[10px] uppercase tracking-widest opacity-50">Anno Riferimento</label>
                <select 
                  value={selectedChannelYear} 
                  onChange={(e) => setSelectedChannelYear(Number(e.target.value))}
                  className="bg-white border border-[#141414] px-4 py-2 text-sm focus:outline-none min-w-[120px]"
                >
                  {availableChannelYears.map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              </div>
              {channelAnalysis && (
                <button 
                  onClick={handleDownloadAnalysis}
                  className="flex items-center gap-2 bg-[#141414] text-white px-4 py-2 text-[10px] uppercase tracking-widest font-bold hover:bg-gray-800 transition-colors h-[38px]"
                >
                  <Download size={14} /> Scarica Report
                </button>
              )}
              <div className="flex flex-col gap-1 relative">
                <label className="text-[10px] uppercase tracking-widest opacity-50">Canali Primari ({selectedChannels.length})</label>
                <div className="relative group">
                  <button 
                    onClick={() => setShowChannelDropdown(!showChannelDropdown)}
                    className="bg-white border border-[#141414] px-4 py-2 text-sm focus:outline-none min-w-[250px] text-left flex justify-between items-center"
                  >
                    <span className="truncate max-w-[200px]">
                      {selectedChannels.length === 0 
                        ? "Seleziona canali..." 
                        : selectedChannels.length === availableChannels.length 
                          ? "Tutti i canali" 
                          : selectedChannels.join(", ")}
                    </span>
                    <ChevronDown size={14} className={cn("transition-transform", showChannelDropdown && "rotate-180")} />
                  </button>
                  
                  {showChannelDropdown && (
                    <>
                      <div 
                        className="fixed inset-0 z-10" 
                        onClick={() => setShowChannelDropdown(false)}
                      />
                      <div className="absolute top-full left-0 w-full bg-white border border-[#141414] mt-1 z-20 shadow-xl max-h-[300px] overflow-y-auto">
                        <div className="p-2 border-b border-gray-100 flex gap-2 sticky top-0 bg-white">
                          <button 
                            onClick={() => setSelectedChannels(availableChannels)}
                            className="text-[9px] uppercase tracking-widest font-bold hover:underline"
                          >
                            Tutti
                          </button>
                          <button 
                            onClick={() => setSelectedChannels([])}
                            className="text-[9px] uppercase tracking-widest font-bold hover:underline"
                          >
                            Nessuno
                          </button>
                        </div>
                        <div className="p-1">
                          {availableChannels.map(c => (
                            <label key={c} className="flex items-center gap-2 p-2 hover:bg-gray-50 cursor-pointer text-xs">
                              <input 
                                type="checkbox"
                                checked={selectedChannels.includes(c)}
                                onChange={() => {
                                  setSelectedChannels(prev => 
                                    prev.includes(c) ? prev.filter(item => item !== c) : [...prev, c]
                                  );
                                }}
                                className="accent-[#141414]"
                              />
                              {c}
                            </label>
                          ))}
                        </div>
                      </div>
                    </>
                  )}
                </div>
              </div>
            </div>
          </div>

          {channelAnalysis ? (
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* 1. Anagrafiche */}
              <div className="border border-[#141414]">
                <div className="bg-yellow-400 p-3 border-b border-[#141414] font-bold text-xs uppercase tracking-widest">
                  Anagrafiche
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[10px] border-collapse">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#141414]">
                        <th className="p-2 border-r border-[#141414]">Store</th>
                        {channelAnalysis.anagrafiche.months.map(m => (
                          <th key={m} className="p-2 border-r border-[#141414] text-center">{MONTH_NAMES[m - 1]}</th>
                        ))}
                        <th className="p-2 text-center font-bold">Totale</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries(channelAnalysis.anagrafiche.data)
                        .sort((a, b) => a[0].localeCompare(b[0]))
                        .map(([store, months]) => {
                          const total = Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0);
                          return (
                            <tr key={store} className="border-b border-gray-100">
                              <td className="p-2 border-r border-[#141414] font-bold">{store}</td>
                              {channelAnalysis.anagrafiche.months.map(m => {
                                const count = (months as Record<number, number>)[m] || 0;
                                return (
                                  <td key={m} className="p-2 border-r border-[#141414] text-center">
                                    {count > 0 ? (
                                      <button 
                                        onClick={() => {
                                          const data = channelAnalysis.filteredPatients.filter(p => (p.store || 'N/D') === store && (p.month + 1) === m);
                                          setChannelDrillDown({
                                            title: `Anagrafiche - ${store}`,
                                            subtitle: `${MONTH_NAMES[m - 1]} - ${selectedChannels.join(', ')}`,
                                            data,
                                            type: 'patients'
                                          });
                                        }}
                                        className="hover:underline text-blue-600 font-bold"
                                      >
                                        {count}
                                      </button>
                                    ) : 0}
                                  </td>
                                );
                              })}
                              <td className="p-2 text-center font-bold bg-gray-50">
                                {total > 0 ? (
                                  <button 
                                    onClick={() => {
                                      const data = channelAnalysis.filteredPatients.filter(p => (p.store || 'N/D') === store);
                                      setChannelDrillDown({
                                        title: `Anagrafiche - ${store}`,
                                        subtitle: `Tutti i mesi - ${selectedChannels.join(', ')}`,
                                        data,
                                        type: 'patients'
                                      });
                                    }}
                                    className="hover:underline text-blue-600 font-bold"
                                  >
                                    {total}
                                  </button>
                                ) : 0}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                    <tfoot className="bg-gray-100 font-bold border-t border-[#141414]">
                      <tr>
                        <td className="p-2 border-r border-[#141414]">Totale complessivo</td>
                        {channelAnalysis.anagrafiche.months.map(m => {
                          const monthTotal = Object.values(channelAnalysis.anagrafiche.data).reduce((sum, months) => 
                            (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
                          return (
                            <td key={m} className="p-2 border-r border-[#141414] text-center">
                              {(monthTotal as number) > 0 ? (
                                <button 
                                  onClick={() => {
                                    const data = (channelAnalysis as any).filteredPatients.filter((p: any) => (p.month + 1) === m);
                                    setChannelDrillDown({
                                      title: `Anagrafiche - Tutte le sedi`,
                                      subtitle: `Mese ${m} - ${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'patients'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {monthTotal as number}
                                </button>
                              ) : 0}
                            </td>
                          );
                        })}
                        <td className="p-2 text-center">
                          {(() => {
                            const grandTotal = Object.values(channelAnalysis.anagrafiche.data).reduce((sum, months) => 
                              (sum as number) + Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0), 0);
                            return (grandTotal as number) > 0 ? (
                              <button 
                                onClick={() => {
                                  setChannelDrillDown({
                                    title: `Anagrafiche - Totale Complessivo`,
                                    subtitle: `Tutte le sedi, tutti i mesi - ${selectedChannels.join(', ')}`,
                                    data: (channelAnalysis as any).filteredPatients,
                                    type: 'patients'
                                  });
                                }}
                                className="hover:underline text-blue-600 font-bold"
                              >
                                {grandTotal as number}
                              </button>
                            ) : 0;
                          })()}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>

              {/* 2. Telefonate */}
              <div className="border border-[#141414]">
                <div className="bg-yellow-400 p-3 border-b border-[#141414] font-bold text-xs uppercase tracking-widest">
                  Telefonate
                </div>
                <div className="p-4">
                  <table className="w-full text-[10px] border-collapse">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#141414]">
                        <th className="p-2 text-left">Esito</th>
                        <th className="p-2 text-right">Conteggio</th>
                        <th className="p-2 text-right">%</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const calls = channelAnalysis.calls as Record<string, number>;
                        const total = Object.values(calls).reduce((a, b) => a + b, 0);
                        return Object.entries(calls)
                          .sort((a, b) => b[1] - a[1])
                          .map(([outcome, count]) => (
                            <tr key={outcome} className="border-b border-gray-100">
                              <td className="p-2">{outcome}</td>
                              <td className="p-2 text-right font-mono">
                                <button 
                                  onClick={() => {
                                    const data = channelAnalysis.filteredCalls.filter(c => (c.contatto || 'N/D') === outcome);
                                    setChannelDrillDown({
                                      title: `Telefonate - ${outcome}`,
                                      subtitle: `${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'calls'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {count}
                                </button>
                              </td>
                              <td className="p-2 text-right font-mono opacity-60">
                                {total > 0 ? ((count / total) * 100).toFixed(0) : 0}%
                              </td>
                            </tr>
                          ));
                      })()}
                    </tbody>
                    <tfoot className="bg-gray-100 font-bold border-t border-[#141414]">
                      <tr>
                        <td className="p-2">Totale complessivo</td>
                        <td className="p-2 text-right font-mono">
                          <button 
                            onClick={() => {
                              setChannelDrillDown({
                                title: `Telefonate - Totale Complessivo`,
                                subtitle: `${selectedChannels.join(', ')}`,
                                data: channelAnalysis.filteredCalls,
                                type: 'calls'
                              });
                            }}
                            className="hover:underline text-blue-600 font-bold"
                          >
                            {Object.values(channelAnalysis.calls as Record<string, number>).reduce((a, b) => a + b, 0)}
                          </button>
                        </td>
                        <td className="p-2 text-right font-mono">100%</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>

              {/* 3. Appuntamenti */}
              <div className="border border-[#141414]">
                <div className="bg-yellow-400 p-3 border-b border-[#141414] font-bold text-xs uppercase tracking-widest">
                  Appuntamenti
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[10px] border-collapse">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#141414]">
                        <th className="p-2 border-r border-[#141414] text-left">Stato</th>
                        {(channelAnalysis.appointments as any).months.map((m: number) => (
                          <th key={m} className="p-2 border-r border-[#141414] text-center">{MONTH_NAMES[m - 1]}</th>
                        ))}
                        <th className="p-2 text-right font-bold">Totale</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries((channelAnalysis.appointments as any).data)
                        .sort((a: any, b: any) => {
                          const sumA = Object.values(a[1]).reduce((x: any, y: any) => x + y, 0) as number;
                          const sumB = Object.values(b[1]).reduce((x: any, y: any) => x + y, 0) as number;
                          return sumB - sumA;
                        })
                        .map(([status, months]) => {
                          const total = Object.values(months as any).reduce((a: any, b: any) => a + b, 0) as number;
                          return (
                            <tr key={status} className="border-b border-gray-100">
                              <td className="p-2 border-r border-[#141414] font-bold">{status}</td>
                              {(channelAnalysis.appointments as any).months.map((m: number) => {
                                const count = (months as any)[m] || 0;
                                return (
                                  <td key={m} className="p-2 border-r border-[#141414] text-center">
                                    {count > 0 ? (
                                      <button 
                                        onClick={() => {
                                          const data = channelAnalysis.filteredAppointments.filter(a => {
                                            const tipo = String(a.tipo || '').toLowerCase().trim();
                                            return tipo === 'prima visita' && (a.contatto || 'N/D') === status && (a.month + 1) === m;
                                          });
                                          setChannelDrillDown({
                                            title: `Appuntamenti - ${status}`,
                                            subtitle: `${MONTH_NAMES[m - 1]} - Prima Visita - ${selectedChannels.join(', ')}`,
                                            data,
                                            type: 'appointments'
                                          });
                                        }}
                                        className="hover:underline text-blue-600 font-bold"
                                      >
                                        {count}
                                      </button>
                                    ) : 0}
                                  </td>
                                );
                              })}
                              <td className="p-2 text-right font-bold bg-gray-50">
                                {total > 0 ? (
                                  <button 
                                    onClick={() => {
                                      const data = channelAnalysis.filteredAppointments.filter(a => {
                                        const tipo = String(a.tipo || '').toLowerCase().trim();
                                        return tipo === 'prima visita' && (a.contatto || 'N/D') === status;
                                      });
                                      setChannelDrillDown({
                                        title: `Appuntamenti - ${status}`,
                                        subtitle: `Tutti i mesi - Prima Visita - ${selectedChannels.join(', ')}`,
                                        data,
                                        type: 'appointments'
                                      });
                                    }}
                                    className="hover:underline text-blue-600 font-bold"
                                  >
                                    {total}
                                  </button>
                                ) : 0}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                    <tfoot className="bg-gray-100 font-bold border-t border-[#141414]">
                      <tr>
                        <td className="p-2 border-r border-[#141414]">Totale complessivo</td>
                        {(channelAnalysis.appointments as any).months.map((m: number) => {
                          const monthTotal = Object.values((channelAnalysis.appointments as any).data).reduce((sum, months) => 
                            (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
                          return (
                            <td key={m} className="p-2 border-r border-[#141414] text-center">
                              {(monthTotal as number) > 0 ? (
                                <button 
                                  onClick={() => {
                                    const data = channelAnalysis.filteredAppointments.filter(a => {
                                      const tipo = String(a.tipo || '').toLowerCase().trim();
                                      return tipo === 'prima visita' && (a.month + 1) === m;
                                    });
                                    setChannelDrillDown({
                                      title: `Appuntamenti - Totale`,
                                      subtitle: `${MONTH_NAMES[m - 1]} - Prima Visita - ${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'appointments'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {monthTotal as number}
                                </button>
                              ) : 0}
                            </td>
                          );
                        })}
                        <td className="p-2 text-right">
                          {(() => {
                            const grandTotal = Object.values((channelAnalysis.appointments as any).data).reduce((sum, months) => 
                              (sum as number) + Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0), 0);
                            return (grandTotal as number) > 0 ? (
                              <button 
                                onClick={() => {
                                  const data = channelAnalysis.filteredAppointments.filter(a => {
                                    const tipo = String(a.tipo || '').toLowerCase().trim();
                                    return tipo === 'prima visita';
                                  });
                                  setChannelDrillDown({
                                    title: `Appuntamenti - Totale Complessivo`,
                                    subtitle: `Prima Visita - ${selectedChannels.join(', ')}`,
                                    data,
                                    type: 'appointments'
                                  });
                                }}
                                className="hover:underline text-blue-600 font-bold"
                              >
                                {grandTotal as number}
                              </button>
                            ) : 0;
                          })()}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>

              {/* 4. Presentati */}
              <div className="border border-[#141414]">
                <div className="bg-yellow-400 p-3 border-b border-[#141414] font-bold text-xs uppercase tracking-widest">
                  Presentati
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[10px] border-collapse">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#141414]">
                        <th className="p-2 border-r border-[#141414] text-left">Tipo Contatto</th>
                        {(channelAnalysis.presentati as any).months.map((m: number) => (
                          <th key={m} className="p-2 border-r border-[#141414] text-center">{MONTH_NAMES[m - 1]}</th>
                        ))}
                        <th className="p-2 text-right font-bold">Totale</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries((channelAnalysis.presentati as any).data)
                        .sort((a: any, b: any) => {
                          const sumA = Object.values(a[1]).reduce((x: any, y: any) => x + y, 0) as number;
                          const sumB = Object.values(b[1]).reduce((x: any, y: any) => x + y, 0) as number;
                          return sumB - sumA;
                        })
                        .map(([type, months]) => {
                          const total = Object.values(months as any).reduce((a: any, b: any) => a + b, 0) as number;
                          return (
                            <tr key={type} className="border-b border-gray-100">
                              <td className="p-2 border-r border-[#141414] font-bold">{type}</td>
                              {(channelAnalysis.presentati as any).months.map((m: number) => {
                                const count = (months as any)[m] || 0;
                                return (
                                  <td key={m} className="p-2 border-r border-[#141414] text-center">
                                    {count > 0 ? (
                                      <button 
                                        onClick={() => {
                                          const data = channelAnalysis.filteredAppointments.filter(a => {
                                            const tipo = String(a.tipo || '').toLowerCase().trim();
                                            const status = String(a.contatto || '').toLowerCase().trim();
                                            return tipo === 'prima visita' && 
                                                   status === 'si è presentato' &&
                                                   (a.contactType || 'N/D') === type &&
                                                   (a.month + 1) === m;
                                          });
                                          setChannelDrillDown({
                                            title: `Presentati - ${type}`,
                                            subtitle: `${MONTH_NAMES[m - 1]} - Prima Visita, Si è presentato - ${selectedChannels.join(', ')}`,
                                            data,
                                            type: 'appointments'
                                          });
                                        }}
                                        className="hover:underline text-blue-600 font-bold"
                                      >
                                        {count}
                                      </button>
                                    ) : 0}
                                  </td>
                                );
                              })}
                              <td className="p-2 text-right font-bold bg-gray-50">
                                {total > 0 ? (
                                  <button 
                                    onClick={() => {
                                      const data = channelAnalysis.filteredAppointments.filter(a => {
                                        const tipo = String(a.tipo || '').toLowerCase().trim();
                                        const status = String(a.contatto || '').toLowerCase().trim();
                                        return tipo === 'prima visita' && 
                                               status === 'si è presentato' &&
                                               (a.contactType || 'N/D') === type;
                                      });
                                      setChannelDrillDown({
                                        title: `Presentati - ${type}`,
                                        subtitle: `Tutti i mesi - Prima Visita, Si è presentato - ${selectedChannels.join(', ')}`,
                                        data,
                                        type: 'appointments'
                                      });
                                    }}
                                    className="hover:underline text-blue-600 font-bold"
                                  >
                                    {total}
                                  </button>
                                ) : 0}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                    <tfoot className="bg-gray-100 font-bold border-t border-[#141414]">
                      <tr>
                        <td className="p-2 border-r border-[#141414]">Totale complessivo</td>
                        {(channelAnalysis.presentati as any).months.map((m: number) => {
                          const monthTotal = Object.values((channelAnalysis.presentati as any).data).reduce((sum, months) => 
                            (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
                          return (
                            <td key={m} className="p-2 border-r border-[#141414] text-center">
                              {(monthTotal as number) > 0 ? (
                                <button 
                                  onClick={() => {
                                    const data = channelAnalysis.filteredAppointments.filter(a => {
                                      const tipo = String(a.tipo || '').toLowerCase().trim();
                                      const status = String(a.contatto || '').toLowerCase().trim();
                                      return tipo === 'prima visita' && status === 'si è presentato' && (a.month + 1) === m;
                                    });
                                    setChannelDrillDown({
                                      title: `Presentati - Totale`,
                                      subtitle: `${MONTH_NAMES[m - 1]} - Prima Visita, Si è presentato - ${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'appointments'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {monthTotal as number}
                                </button>
                              ) : 0}
                            </td>
                          );
                        })}
                        <td className="p-2 text-right">
                          {(() => {
                            const grandTotal = Object.values((channelAnalysis.presentati as any).data).reduce((sum, months) => 
                              (sum as number) + Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0), 0);
                            return (grandTotal as number) > 0 ? (
                              <button 
                                onClick={() => {
                                  const data = channelAnalysis.filteredAppointments.filter(a => {
                                    const tipo = String(a.tipo || '').toLowerCase().trim();
                                    const status = String(a.contatto || '').toLowerCase().trim();
                                    return tipo === 'prima visita' && status === 'si è presentato';
                                  });
                                  setChannelDrillDown({
                                    title: `Presentati - Totale Complessivo`,
                                    subtitle: `Prima Visita, Si è presentato - ${selectedChannels.join(', ')}`,
                                    data,
                                    type: 'appointments'
                                  });
                                }}
                                className="hover:underline text-blue-600 font-bold"
                              >
                                {grandTotal as number}
                              </button>
                            ) : 0;
                          })()}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>

              {/* 5. Prove */}
              <div className="border border-[#141414]">
                <div className="bg-yellow-400 p-3 border-b border-[#141414] font-bold text-xs uppercase tracking-widest">
                  Prove
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[10px] border-collapse">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#141414]">
                        <th className="p-2 border-r border-[#141414] text-left">Stato documento</th>
                        {(channelAnalysis.trials as any).months.map((m: number) => (
                          <th key={m} className="p-2 border-r border-[#141414] text-center">{MONTH_NAMES[m - 1]}</th>
                        ))}
                        <th className="p-2 text-right font-bold">Totale</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries((channelAnalysis.trials as any).data)
                        .sort((a: any, b: any) => {
                          const sumA = Object.values(a[1]).reduce((x: any, y: any) => x + y, 0) as number;
                          const sumB = Object.values(b[1]).reduce((x: any, y: any) => x + y, 0) as number;
                          return sumB - sumA;
                        })
                        .map(([status, months]) => {
                          const total = Object.values(months as any).reduce((a: any, b: any) => a + b, 0) as number;
                          return (
                            <tr key={status} className="border-b border-gray-100">
                              <td className="p-2 border-r border-[#141414] font-bold">{status}</td>
                              {(channelAnalysis.trials as any).months.map((m: number) => {
                                const count = (months as any)[m] || 0;
                                return (
                                  <td key={m} className="p-2 border-r border-[#141414] text-center">
                                    {count > 0 ? (
                                      <button 
                                        onClick={() => {
                                          const data = channelAnalysis.filteredTrials.filter(t => (t.statoDocumento || 'N/D') === status && (t.month + 1) === m);
                                          setChannelDrillDown({
                                            title: `Prove - ${status}`,
                                            subtitle: `${MONTH_NAMES[m - 1]} - ${selectedChannels.join(', ')}`,
                                            data,
                                            type: 'trials'
                                          });
                                        }}
                                        className="hover:underline text-blue-600 font-bold"
                                      >
                                        {count}
                                      </button>
                                    ) : 0}
                                  </td>
                                );
                              })}
                              <td className="p-2 text-right font-bold bg-gray-50">
                                {total > 0 ? (
                                  <button 
                                    onClick={() => {
                                      const data = channelAnalysis.filteredTrials.filter(t => (t.statoDocumento || 'N/D') === status);
                                      setChannelDrillDown({
                                        title: `Prove - ${status}`,
                                        subtitle: `Tutti i mesi - ${selectedChannels.join(', ')}`,
                                        data,
                                        type: 'trials'
                                      });
                                    }}
                                    className="hover:underline text-blue-600 font-bold"
                                  >
                                    {total}
                                  </button>
                                ) : 0}
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                    <tfoot className="bg-gray-100 font-bold border-t border-[#141414]">
                      <tr>
                        <td className="p-2 border-r border-[#141414]">Totale complessivo</td>
                        {(channelAnalysis.trials as any).months.map((m: number) => {
                          const monthTotal = Object.values((channelAnalysis.trials as any).data).reduce((sum, months) => 
                            (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
                          return (
                            <td key={m} className="p-2 border-r border-[#141414] text-center">
                              {(monthTotal as number) > 0 ? (
                                <button 
                                  onClick={() => {
                                    const data = channelAnalysis.filteredTrials.filter(t => (t.month + 1) === m);
                                    setChannelDrillDown({
                                      title: `Prove - Totale`,
                                      subtitle: `${MONTH_NAMES[m - 1]} - ${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'trials'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {monthTotal as number}
                                </button>
                              ) : 0}
                            </td>
                          );
                        })}
                        <td className="p-2 text-right">
                          {(() => {
                            const grandTotal = Object.values((channelAnalysis.trials as any).data).reduce((sum, months) => 
                              (sum as number) + Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0), 0);
                            return (grandTotal as number) > 0 ? (
                              <button 
                                onClick={() => {
                                  setChannelDrillDown({
                                    title: `Prove - Totale Complessivo`,
                                    subtitle: `${selectedChannels.join(', ')}`,
                                    data: channelAnalysis.filteredTrials,
                                    type: 'trials'
                                  });
                                }}
                                className="hover:underline text-blue-600 font-bold"
                              >
                                {grandTotal as number}
                              </button>
                            ) : 0;
                          })()}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>

              {/* 6. Fatturato */}
              <div className="border border-[#141414]">
                <div className="bg-yellow-400 p-3 border-b border-[#141414] font-bold text-xs uppercase tracking-widest">
                  Fatturato
                </div>
                <div className="p-4 overflow-x-auto">
                  <table className="w-full text-[10px] border-collapse">
                    <thead>
                      <tr className="bg-gray-50 border-b border-[#141414]">
                        <th className="p-2 text-left border-r border-[#141414]">Audioprotesista</th>
                        {(channelAnalysis.sales as any).months.map((m: number) => (
                          <th key={m} className="p-2 text-center border-r border-[#141414]">{MONTH_NAMES[m - 1]}</th>
                        ))}
                        <th className="p-2 text-right">Totale</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries((channelAnalysis.sales as any).data)
                        .sort((a, b) => {
                          const sumA = Object.values(a[1] as Record<number, number>).reduce((x, y) => x + y, 0);
                          const sumB = Object.values(b[1] as Record<number, number>).reduce((x, y) => x + y, 0);
                          return sumB - sumA;
                        })
                        .map(([person, months]) => {
                          const rowTotal = Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0);
                          return (
                            <tr key={person} className="border-b border-gray-100">
                              <td className="p-2 font-bold border-r border-[#141414]">{person}</td>
                              {(channelAnalysis.sales as any).months.map((m: number) => {
                                const val = (months as Record<number, number>)[m] || 0;
                                return (
                                  <td key={m} className="p-2 border-r border-[#141414] text-center font-mono">
                                    {val > 0 ? (
                                      <button 
                                        onClick={() => {
                                          const data = channelAnalysis.filteredSales.filter(s => 
                                            (s.audioprotesista || 'N/D') === person && (s.month + 1) === m
                                          );
                                          setChannelDrillDown({
                                            title: `Fatturato - ${person}`,
                                            subtitle: `${MONTH_NAMES[m - 1]} - ${selectedChannels.join(', ')}`,
                                            data,
                                            type: 'sales'
                                          });
                                        }}
                                        className="hover:underline text-blue-600 font-bold"
                                      >
                                        {new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' }).format(val)}
                                      </button>
                                    ) : '-'}
                                  </td>
                                );
                              })}
                              <td className="p-2 text-right font-mono font-bold">
                                <button 
                                  onClick={() => {
                                    const data = channelAnalysis.filteredSales.filter(s => (s.audioprotesista || 'N/D') === person);
                                    setChannelDrillDown({
                                      title: `Fatturato - ${person}`,
                                      subtitle: `Tutti i mesi - ${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'sales'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' }).format(rowTotal)}
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                    <tfoot className="bg-gray-100 font-bold border-t border-[#141414]">
                      <tr>
                        <td className="p-2 border-r border-[#141414]">Totale complessivo</td>
                        {(channelAnalysis.sales as any).months.map((m: number) => {
                          const monthTotal = Object.values((channelAnalysis.sales as any).data).reduce((sum, months) => 
                            (sum as number) + (((months as Record<number, number>)[m]) || 0), 0);
                          return (
                            <td key={m} className="p-2 border-r border-[#141414] text-center font-mono">
                              {(monthTotal as number) > 0 ? (
                                <button 
                                  onClick={() => {
                                    const data = channelAnalysis.filteredSales.filter(s => (s.month + 1) === m);
                                    setChannelDrillDown({
                                      title: `Fatturato - Tutti gli audioprotesisti`,
                                      subtitle: `${MONTH_NAMES[m - 1]} - ${selectedChannels.join(', ')}`,
                                      data,
                                      type: 'sales'
                                    });
                                  }}
                                  className="hover:underline text-blue-600 font-bold"
                                >
                                  {new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' }).format(monthTotal as number)}
                                </button>
                              ) : '-'}
                            </td>
                          );
                        })}
                        <td className="p-2 text-right font-mono">
                          <button 
                            onClick={() => {
                              setChannelDrillDown({
                                title: `Fatturato - Totale Complessivo`,
                                subtitle: `${selectedChannels.join(', ')}`,
                                data: channelAnalysis.filteredSales,
                                type: 'sales'
                              });
                            }}
                            className="hover:underline text-blue-600 font-bold"
                          >
                            {new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' }).format(
                              Object.values((channelAnalysis.sales as any).data).reduce((sum, months) => 
                                (sum as number) + Object.values(months as Record<number, number>).reduce((a, b) => a + b, 0), 0) as number
                            )}
                          </button>
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
            </div>
          ) : (
            <div className="p-20 text-center border border-dashed border-gray-300 rounded-lg">
              <div className="text-4xl mb-4">📊</div>
              <div className="font-serif italic text-xl opacity-40">Seleziona un canale primario per visualizzare l'analisi dettagliata.</div>
            </div>
          )}
        </div>
      ) : activeTab === 'noah' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <FileCode size={32} /> Analizza Noah
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Carica uno o più file XML esportati da Noah per estrarre i pazienti che hanno un numero di telefono valido e un esame audiometrico con StimulusFrequency 1000Hz e StimulusLevel {'>'} {noahThreshold}dB.
              </p>
            </div>
            <div className="flex items-center gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-[10px] uppercase tracking-widest opacity-50">Soglia 1000Hz (dB)</label>
                <select 
                  value={noahThreshold} 
                  onChange={(e) => setNoahThreshold(Number(e.target.value))}
                  className="bg-white border border-[#141414] px-4 py-2 text-sm focus:outline-none"
                  disabled={loading.noah}
                >
                  {[35, 40, 45, 50, 55, 60].map(val => (
                    <option key={val} value={val}>{val} dB</option>
                  ))}
                </select>
              </div>
              <label className={cn(
                "group cursor-pointer shrink-0",
                loading.noah && "pointer-events-none opacity-50"
              )}>
                <div className="bg-[#141414] text-white px-8 py-4 flex items-center justify-center gap-3 hover:bg-[#333] transition-all active:scale-95 shadow-lg">
                  {loading.noah ? (
                    <Loader2 size={20} className="animate-spin" />
                  ) : (
                    <Upload size={20} />
                  )}
                  <span className="text-xs uppercase tracking-[0.2em] font-bold">
                    {loading.noah ? "Elaborazione..." : "Seleziona File XML"}
                  </span>
                </div>
                <input 
                  type="file" 
                  accept=".xml" 
                  multiple
                  className="hidden" 
                  onChange={handleNoahUpload}
                  disabled={loading.noah}
                />
              </label>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
            <div className="border border-[#141414] p-4 bg-gray-50">
              <div className="text-[10px] uppercase tracking-widest opacity-50 mb-2">Criterio 1</div>
              <div className="font-bold flex items-center gap-2">
                <CheckCircle2 size={16} className="text-emerald-600" />
                Presenza Numero di Telefono (escluso "0")
              </div>
            </div>
            <div className="border border-[#141414] p-4 bg-gray-50">
              <div className="text-[10px] uppercase tracking-widest opacity-50 mb-2">Criterio 2</div>
              <div className="font-bold flex items-center gap-2">
                <CheckCircle2 size={16} className="text-emerald-600" />
                Audiometria (1000Hz {'>'} {noahThreshold}dB)
              </div>
            </div>
          </div>

          {noahPatients.length > 0 && (
            <div className="mb-6 flex justify-end">
              <button
                onClick={handleDownloadNoahExcel}
                className="bg-emerald-600 text-white px-6 py-3 flex items-center gap-2 hover:bg-emerald-700 transition-all active:scale-95 shadow-md text-xs uppercase tracking-widest font-bold"
              >
                <Download size={18} />
                Scarica Excel ({noahPatients.length})
              </button>
            </div>
          )}

          {error && (
            <div className="bg-red-50 border border-red-200 text-red-700 p-4 mb-8 flex items-center gap-3">
              <AlertCircle size={20} />
              <span className="text-sm font-medium">{error}</span>
            </div>
          )}

          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse border border-[#141414]">
              <thead>
                <tr className="bg-[#141414] text-white">
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Cognome</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Nome</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Telefono</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">Valore 1000Hz</th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold">Data Creazione</th>
                </tr>
              </thead>
              <tbody className="font-mono text-xs">
                {noahPatients.length > 0 ? (
                  noahPatients.map((p, idx) => (
                    <tr key={idx} className="border-b border-[#141414] hover:bg-gray-50 transition-colors">
                      <td className="p-4 border-r border-[#141414] font-bold">{p.cognome}</td>
                      <td className="p-4 border-r border-[#141414]">{p.nome}</td>
                      <td className="p-4 border-r border-[#141414] font-bold text-emerald-700">{p.telefono}</td>
                      <td className="p-4 border-r border-[#141414] font-bold text-blue-700">{p.valore1000 !== null ? `${p.valore1000} dB` : 'N/D'}</td>
                      <td className="p-4 opacity-60">{p.dataCreazione || 'N/D'}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={5} className="p-12 text-center opacity-40 italic font-serif text-lg">
                      Nessun paziente trovato con i criteri specificati. Carica un file XML per iniziare.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      ) : activeTab === 'cap' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <Building2 size={32} /> Strutture Esterne (CAP)
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Associa uno o più CAP alle strutture esterne per facilitare l'assegnazione dei pazienti.
              </p>
            </div>
            <div className="flex items-center gap-4">
              <div className="flex items-center border border-[#141414] bg-white">
                <Search size={16} className="ml-3 opacity-40" />
                <input 
                  type="text"
                  placeholder="Cerca struttura..."
                  value={capSearch}
                  onChange={(e) => {
                    setCapSearch(e.target.value);
                    setCapPage(1);
                  }}
                  className="bg-transparent font-mono font-bold text-xs focus:outline-none px-3 py-3 w-48"
                />
              </div>
              <div className="flex items-center border border-[#141414]">
                <input 
                  type="text"
                  placeholder="Inserisci CAP..."
                  value={capInput}
                  onChange={(e) => setCapInput(e.target.value)}
                  className="bg-transparent font-mono font-bold text-xs focus:outline-none px-4 py-3 w-32"
                />
                <button
                  onClick={handleAssociateCap}
                  disabled={!capInput.trim() || selectedStoreIds.length === 0 || loading.stores}
                  className={cn(
                    "px-6 py-3 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold transition-all",
                    (!capInput.trim() || selectedStoreIds.length === 0 || loading.stores) ? "opacity-30 cursor-not-allowed" : "hover:bg-[#333] active:scale-95"
                  )}
                >
                  Associa
                </button>
              </div>
              <div className="flex items-center gap-2 text-[10px] uppercase tracking-widest opacity-60 font-mono">
                {filteredStores.length} strutture trovate
              </div>
            </div>
          </div>

          <div className="overflow-x-auto mb-8">
            <table className="w-full text-left border-collapse border border-[#141414]">
              <thead>
                <tr className="bg-[#141414] text-white">
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20 w-12 text-center">
                    <input 
                      type="checkbox" 
                      onChange={(e) => {
                        if (e.target.checked) {
                          const currentPageIds = filteredStores
                            .slice((capPage - 1) * itemsPerPage, capPage * itemsPerPage)
                            .map(s => s.id);
                          setSelectedStoreIds(prev => Array.from(new Set([...prev, ...currentPageIds])));
                        } else {
                          const currentPageIds = filteredStores
                            .slice((capPage - 1) * itemsPerPage, capPage * itemsPerPage)
                            .map(s => s.id);
                          setSelectedStoreIds(prev => prev.filter(id => !currentPageIds.includes(id)));
                        }
                      }}
                      checked={filteredStores.length > 0 && filteredStores.slice((capPage - 1) * itemsPerPage, capPage * itemsPerPage).every(s => selectedStoreIds.includes(s.id))}
                    />
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">
                    <button 
                      onClick={() => setCapSortConfig({ key: 'nome', direction: capSortConfig?.key === 'nome' && capSortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                      className="flex items-center gap-2 hover:opacity-70 transition-opacity"
                    >
                      Nome Struttura
                      {capSortConfig?.key === 'nome' && (
                        capSortConfig.direction === 'asc' ? <ChevronRight size={12} className="-rotate-90" /> : <ChevronRight size={12} className="rotate-90" />
                      )}
                    </button>
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">
                    <button 
                      onClick={() => setCapSortConfig({ key: 'tipo', direction: capSortConfig?.key === 'tipo' && capSortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                      className="flex items-center gap-2 hover:opacity-70 transition-opacity"
                    >
                      Tipo
                      {capSortConfig?.key === 'tipo' && (
                        capSortConfig.direction === 'asc' ? <ChevronRight size={12} className="-rotate-90" /> : <ChevronRight size={12} className="rotate-90" />
                      )}
                    </button>
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">
                    <button 
                      onClick={() => setCapSortConfig({ key: 'indirizzo', direction: capSortConfig?.key === 'indirizzo' && capSortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                      className="flex items-center gap-2 hover:opacity-70 transition-opacity"
                    >
                      Indirizzo
                      {capSortConfig?.key === 'indirizzo' && (
                        capSortConfig.direction === 'asc' ? <ChevronRight size={12} className="-rotate-90" /> : <ChevronRight size={12} className="rotate-90" />
                      )}
                    </button>
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">
                    <button 
                      onClick={() => setCapSortConfig({ key: 'citta', direction: capSortConfig?.key === 'citta' && capSortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                      className="flex items-center gap-2 hover:opacity-70 transition-opacity"
                    >
                      Città
                      {capSortConfig?.key === 'citta' && (
                        capSortConfig.direction === 'asc' ? <ChevronRight size={12} className="-rotate-90" /> : <ChevronRight size={12} className="rotate-90" />
                      )}
                    </button>
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">
                    <button 
                      onClick={() => setCapSortConfig({ key: 'deposito', direction: capSortConfig?.key === 'deposito' && capSortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                      className="flex items-center gap-2 hover:opacity-70 transition-opacity"
                    >
                      Deposito
                      {capSortConfig?.key === 'deposito' && (
                        capSortConfig.direction === 'asc' ? <ChevronRight size={12} className="-rotate-90" /> : <ChevronRight size={12} className="rotate-90" />
                      )}
                    </button>
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold border-r border-white/20">
                    <button 
                      onClick={() => setCapSortConfig({ key: 'createdAt', direction: capSortConfig?.key === 'createdAt' && capSortConfig.direction === 'asc' ? 'desc' : 'asc' })}
                      className="flex items-center gap-2 hover:opacity-70 transition-opacity"
                    >
                      Creato il
                      {capSortConfig?.key === 'createdAt' && (
                        capSortConfig.direction === 'asc' ? <ChevronRight size={12} className="-rotate-90" /> : <ChevronRight size={12} className="rotate-90" />
                      )}
                    </button>
                  </th>
                  <th className="p-4 text-[10px] uppercase tracking-widest font-bold">CAP Associati</th>
                </tr>
              </thead>
              <tbody className="font-mono text-xs">
                {groupedStores.length > 0 ? (
                  groupedStores.map(([deposito, items]) => (
                    <React.Fragment key={deposito}>
                      <tr className="bg-gray-100/50">
                        <td colSpan={8} className="p-2 pl-4 text-[10px] font-bold uppercase tracking-widest text-[#141414] border-b border-[#141414]">
                          <div className="flex items-center gap-2">
                            <Warehouse size={12} /> {deposito}
                          </div>
                        </td>
                      </tr>
                      {items.map((s, idx) => {
                        const associatedCaps = storeCaps
                          .filter(sc => sc.storeId === s.id)
                          .map(sc => sc.cap);
                        
                        return (
                          <tr key={s.id || idx} className={cn(
                            "border-b border-[#141414] hover:bg-gray-50 transition-colors",
                            selectedStoreIds.includes(s.id) ? "bg-emerald-50/30" : ""
                          )}>
                            <td className="p-4 border-r border-[#141414] text-center">
                              <input 
                                type="checkbox" 
                                checked={selectedStoreIds.includes(s.id)}
                                onChange={() => toggleStoreSelection(s.id)}
                              />
                            </td>
                            <td className="p-4 border-r border-[#141414] font-bold">{s.nome}</td>
                            <td className="p-4 border-r border-[#141414] opacity-70">{s.tipo || 'N/D'}</td>
                            <td className="p-4 border-r border-[#141414]">{s.indirizzo}</td>
                            <td className="p-4 border-r border-[#141414]">{s.citta}</td>
                            <td className="p-4 border-r border-[#141414] font-bold text-gray-500">{s.deposito || 'N/D'}</td>
                            <td className="p-4 border-r border-[#141414] opacity-60">
                              {s.createdAt ? new Date(s.createdAt).toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : 'N/D'}
                            </td>
                            <td className="p-4">
                              <div className="flex flex-wrap gap-1">
                                {s.cap && (
                                  <span className="bg-gray-100 text-gray-800 px-2 py-0.5 rounded text-[9px] font-bold border border-gray-200" title="CAP Struttura">
                                    {s.cap}
                                  </span>
                                )}
                                {associatedCaps.length > 0 ? (
                                  associatedCaps.map((cap, cIdx) => {
                                    const association = storeCaps.find(sc => sc.storeId === s.id && sc.cap === cap);
                                    return (
                                      <span 
                                        key={cIdx} 
                                        onClick={async () => {
                                          if (association?.id) {
                                            try {
                                              await deleteDoc(doc(db, 'storeCaps', association.id));
                                              setStoreCaps(prev => prev.filter(sc => sc.id !== association.id));
                                            } catch (err) {
                                              console.error("Errore nella rimozione del CAP:", err);
                                            }
                                          }
                                        }}
                                        className="bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded text-[9px] font-bold border border-emerald-200 cursor-pointer hover:bg-red-100 hover:text-red-800 hover:border-red-200 transition-colors"
                                        title="Clicca per rimuovere"
                                      >
                                        {cap}
                                      </span>
                                    );
                                  })
                                ) : !s.cap ? (
                                  <span className="opacity-30 italic">Nessun CAP</span>
                                ) : null}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                  ))
                ) : (
                  <tr>
                    <td colSpan={8} className="p-12 text-center opacity-40 italic font-serif text-lg">
                      Nessuna struttura trovata. Carica il file store.xlsx nella sezione "Gestione Pazienti".
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      ) : activeTab === 'budget' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <TrendingUp size={32} /> Gestione Budget
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Imposta la distribuzione mensile del fatturato e il budget annuale per ogni audioprotesista.
              </p>
            </div>
            <div className="flex items-center gap-3">
              <button
                onClick={() => setIsBudgetSummaryOpen(true)}
                className="bg-[#141414] text-white px-6 py-3 text-[10px] uppercase tracking-[0.2em] font-bold hover:bg-[#333] transition-all flex items-center gap-2"
              >
                <FileSpreadsheet size={16} /> Situazione
              </button>
              {isSavingBudget && (
              <div className="flex items-center gap-2 text-emerald-600 font-bold text-xs uppercase tracking-widest animate-pulse">
                <Loader2 size={16} className="animate-spin" /> Salvataggio in corso...
              </div>
            )}
          </div>
        </div>

          <div className="space-y-12">
            {/* Monthly Percentages Section */}
            <div className="bg-gray-50 border border-[#141414] p-6">
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-[10px] uppercase tracking-[0.2em] font-bold flex items-center gap-2">
                  1. Distribuzione Mensile Fatturato (%)
                  <span className={cn(
                    "ml-4 px-2 py-0.5 rounded text-white",
                    Math.abs(monthlyDistribution.reduce((a, b) => a + b, 0) - 100) < 0.1 ? "bg-emerald-600" : "bg-red-500"
                  )}>
                    Totale: {monthlyDistribution.reduce((a, b) => a + b, 0).toFixed(1)}%
                  </span>
                </h3>
                <button 
                  onClick={() => {
                    const equal = new Array(12).fill(100/12);
                    setMonthlyDistribution(equal);
                    saveMonthlyDistribution(equal);
                  }}
                  className="text-[9px] uppercase tracking-widest opacity-50 hover:opacity-100 transition-all font-bold"
                >
                  [ Reimposta Pari 8.3% ]
                </button>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-6 xl:grid-cols-12 gap-4">
                {MONTHS.map((m, idx) => (
                  <div key={m} className="flex flex-col gap-1">
                    <label className="text-[9px] uppercase tracking-widest opacity-50 font-bold">{m.slice(0, 3)}</label>
                    <div className="relative">
                      <input 
                        type="number"
                        step="0.1"
                        value={monthlyDistribution[idx]}
                        onChange={(e) => {
                          const val = parseFloat(e.target.value) || 0;
                          const newDist = [...monthlyDistribution];
                          newDist[idx] = val;
                          setMonthlyDistribution(newDist);
                        }}
                        onBlur={() => saveMonthlyDistribution(monthlyDistribution)}
                        className="w-full bg-white border border-[#141414] px-3 py-2 text-sm font-mono focus:outline-none focus:ring-1 focus:ring-[#141414]"
                      />
                      <span className="absolute right-2 top-1/2 -translate-y-1/2 text-[10px] opacity-30">%</span>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Personnel Budgets Section */}
            <div className="border border-[#141414]">
              <div className="bg-[#141414] text-white p-3 font-bold text-[10px] uppercase tracking-widest">
                2. Budget Annuale per Audioprotesista
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse font-mono text-[11px]">
                  <thead>
                    <tr className="bg-gray-100 border-b border-[#141414]">
                      <th className="p-4 border-r border-[#141414] w-[350px]">Audioprotesista</th>
                      <th className="p-4 border-r border-[#141414] w-[350px] bg-white">Budget Annuale</th>
                      {MONTHS.map(m => (
                        <th key={m} className="p-4 border-r border-gray-200 text-center last:border-r-0 opacity-50 w-[140px]">
                          {m.slice(0, 3)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {/* Totals Row At Top */}
                    {orderedPersonnelList.length > 0 && (
                      <tr className="bg-[#141414] text-white font-bold border-b-2 border-[#141414]">
                        <td className="p-4 border-r border-white/10 uppercase tracking-widest text-[10px] bg-white/5">TOTALE BUDGET CONSOLIDATO</td>
                        <td className="p-4 border-r border-white/10 bg-white/10 text-emerald-400 text-right text-2xl font-mono">
                          {personnelBudgets.reduce((acc, b) => acc + b.annualBudget, 0).toLocaleString('it-IT')}
                        </td>
                        {monthlyDistribution.map((pct, mIdx) => {
                          const totalAnnual = personnelBudgets.reduce((acc, b) => acc + b.annualBudget, 0);
                          const totalMonth = (totalAnnual * pct) / 100;
                          return (
                            <td key={mIdx} className="p-4 text-center border-r border-white/5 last:border-r-0 opacity-80 text-xs">
                              {totalMonth > 0 ? Math.round(totalMonth).toLocaleString('it-IT') : '-'}
                            </td>
                          );
                        })}
                      </tr>
                    )}
                    {orderedPersonnelList.length > 0 ? (
                      orderedPersonnelList.map((name, idx) => {
                        const personBudget = personnelBudgets.find(b => b.name === name);
                        const annualVal = personBudget?.annualBudget || 0;
                        
                        return (
                          <tr key={name} className="border-b border-gray-100 hover:bg-gray-50 transition-colors group">
                            <td className="p-4 border-r border-[#141414] font-bold">
                              <div className="flex items-center justify-between gap-4">
                                <span className="truncate flex-1">{name}</span>
                                <div className="flex flex-col gap-1 shrink-0">
                                  <button 
                                    onClick={() => movePersonnel(idx, 'up')}
                                    disabled={idx === 0}
                                    className="p-1.5 bg-white border border-[#141414] hover:bg-gray-100 rounded-sm disabled:opacity-20 cursor-pointer shadow-sm"
                                    title="Sposta su"
                                  >
                                    <ChevronUp size={14} />
                                  </button>
                                  <button 
                                    onClick={() => movePersonnel(idx, 'down')}
                                    disabled={idx === orderedPersonnelList.length - 1}
                                    className="p-1.5 bg-white border border-[#141414] hover:bg-gray-100 rounded-sm disabled:opacity-20 cursor-pointer shadow-sm"
                                    title="Sposta giù"
                                  >
                                    <ChevronDown size={14} />
                                  </button>
                                </div>
                              </div>
                            </td>
                             <td className="p-0 border-r border-[#141414] bg-white relative">
                              <div className="flex items-center w-full h-full">
                                <input 
                                  type="text"
                                  value={(annualVal || 0).toLocaleString('it-IT')}
                                  placeholder="0"
                                  onChange={(e) => {
                                    const rawValue = e.target.value.replace(/\./g, '').replace(/[^0-9]/g, '');
                                    const val = parseInt(rawValue) || 0;
                                    setPersonnelBudgets(prev => {
                                      const existing = prev.find(p => p.name === name);
                                      if (existing) return prev.map(p => p.name === name ? { ...p, annualBudget: val } : p);
                                      return [...prev, { year: selectedYear, name, annualBudget: val }];
                                    });
                                  }}
                                  onBlur={(e) => {
                                    const rawValue = e.target.value.replace(/\./g, '').replace(/[^0-9]/g, '');
                                    savePersonBudget(name, parseInt(rawValue) || 0);
                                  }}
                                  className="w-full h-[70px] bg-transparent px-4 text-xl font-mono font-bold focus:outline-none focus:bg-emerald-50/30 transition-colors text-right"
                                />
                              </div>
                            </td>
                            {monthlyDistribution.map((pct, mIdx) => {
                              const mBudget = (annualVal * pct) / 100;
                              return (
                                <td key={mIdx} className="p-4 border-r border-gray-100 last:border-r-0 text-center opacity-70">
                                  {mBudget > 0 ? Math.round(mBudget).toLocaleString('it-IT') : '-'}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })
                    ) : (
                      <tr>
                        <td colSpan={14} className="p-12 text-center opacity-40 italic font-serif text-lg">
                          Carica dati di vendita per visualizzare l'elenco degli audioprotesisti.
                        </td>
                      </tr>
                    )}
                  {/* Totals already displayed at the top */}
                  </tbody>
                  {/* Footer removed since totals are at top */}
                </table>
              </div>
            </div>
          </div>
        </div>
      ) : activeTab === 'comparisons' ? (
        <div className="bg-white border border-[#141414] p-8">
          <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8 border-b border-[#141414] pb-8">
            <div>
              <h2 className="font-serif italic text-3xl mb-2 flex items-center gap-3">
                <TrendingUp size={32} /> Comparazione 3 Anni
              </h2>
              <p className="text-sm opacity-60 max-w-xl">
                Analisi del fatturato netto (vendite - note di credito) per audioprotesisti e recapiti negli ultimi tre anni.
              </p>
            </div>
            <div className="flex items-center gap-4">
              <div className="flex flex-col gap-1">
                <label className="text-[10px] uppercase tracking-widest opacity-50">Anno di Riferimento</label>
                <select 
                  value={comparisonAnchorYear} 
                  onChange={(e) => setComparisonAnchorYear(Number(e.target.value))}
                  className="bg-white border border-[#141414] px-4 py-2 text-sm focus:outline-none"
                >
                  {availableYears.map(y => (
                    <option key={y} value={y}>{y}</option>
                  ))}
                </select>
              </div>
              <button
                onClick={handleDownloadComparisonExcel}
                className="flex items-center gap-2 px-4 py-2 bg-[#141414] text-white hover:bg-[#333] transition-all text-[10px] uppercase tracking-widest font-bold self-end"
              >
                <Download size={14} /> Download Excel
              </button>
            </div>
          </div>

          {/* Depositi Filter */}
          {comparison3YearData.availableDepositi && comparison3YearData.availableDepositi.length > 0 && (
            <div className="mb-12 p-6 bg-gray-50 border border-[#141414]">
              <div className="flex flex-wrap items-center gap-4">
                <span className="text-[10px] uppercase tracking-widest font-bold opacity-60 flex items-center gap-2 mr-4">
                  <Warehouse size={16} /> Filtra Depositi:
                </span>
                <button
                  onClick={() => setVisibleDepositi(comparison3YearData.availableDepositi)}
                  className="px-4 py-2 bg-white border border-[#141414] text-[9px] uppercase tracking-widest font-bold hover:bg-gray-100 transition-all"
                >
                  Mostra Tutti
                </button>
                <button
                  onClick={() => setVisibleDepositi([])}
                  className="px-4 py-2 bg-white border border-[#141414] text-[9px] uppercase tracking-widest font-bold hover:bg-gray-100 transition-all"
                >
                  Nascondi Tutti
                </button>
                <div className="h-8 w-px bg-[#141414]/10 mx-2" />
                <div className="flex flex-wrap gap-2">
                  {comparison3YearData.availableDepositi.map(dep => (
                    <button
                      key={dep}
                      onClick={() => {
                        setVisibleDepositi(prev => 
                          prev.includes(dep) ? prev.filter(d => d !== dep) : [...prev, dep]
                        );
                      }}
                      className={cn(
                        "px-4 py-2 text-[10px] uppercase tracking-widest font-bold border transition-all",
                        visibleDepositi.includes(dep)
                          ? "bg-[#141414] text-white border-[#141414]"
                          : "bg-white text-[#141414] border-[#141414]/20 hover:border-[#141414]"
                      )}
                    >
                      {dep}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          )}

          <div className="space-y-16">
            {/* Audioprotesisti Section */}
            <div>
              <h3 className="text-xs uppercase tracking-[0.3em] font-bold mb-6 flex items-center gap-2">
                <Users size={16} /> Analisi per Audioprotesista
              </h3>
              <ComparisonTable 
                data={comparison3YearData.audioprotesisti} 
                years={[comparisonAnchorYear, comparisonAnchorYear - 1, comparisonAnchorYear - 2]} 
                onCellClick={handleCellClick}
                comparisonType="audioprotesista"
              />
            </div>

            {/* Recapiti Section */}
            <div>
              <h3 className="text-xs uppercase tracking-[0.3em] font-bold mb-6 flex items-center gap-2">
                <Store size={16} /> Analisi per Struttura (Recapiti)
              </h3>
              <ComparisonTable 
                data={comparison3YearData.recapiti} 
                years={[comparisonAnchorYear, comparisonAnchorYear - 1, comparisonAnchorYear - 2]} 
                onCellClick={handleCellClick}
                comparisonType="recapiti"
              />
            </div>
          </div>
        </div>
      ) : null}
      <footer className="mt-12 pt-8 border-t border-[#141414] flex flex-col md:flex-row justify-between items-center gap-4 opacity-40 text-[10px] uppercase tracking-[0.3em]">
        <span>AskoltaOra CRM System v1.0</span>
        <span>© 2026 Ditte AskoltaOra</span>
      </footer>
      {/* Channel Drill-down Modal */}
      {channelDrillDown && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-[110] flex items-center justify-center p-4">
          <div className="bg-white border border-[#141414] w-full max-w-5xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-gray-50">
              <div>
                <h3 className="font-serif italic text-2xl">{channelDrillDown.title}</h3>
                <p className="text-[10px] uppercase tracking-widest opacity-50">
                  {channelDrillDown.subtitle}
                </p>
              </div>
              <button 
                onClick={() => setChannelDrillDown(null)}
                className="w-10 h-10 flex items-center justify-center border border-[#141414] hover:bg-black hover:text-white transition-all"
              >
                ✕
              </button>
            </div>
            
            <div className="flex-1 overflow-y-auto p-6">
              <table className="w-full text-left border-collapse font-mono text-[10px]">
                <thead>
                  <tr className="border-b border-[#141414] opacity-50 bg-gray-50">
                    <th className="p-2">Data</th>
                    <th className="p-2">Paziente/Contatto</th>
                    <th className="p-2">Canale</th>
                    {channelDrillDown.type === 'patients' && <th className="p-2">Sede</th>}
                    {channelDrillDown.type === 'calls' && <th className="p-2">Esito</th>}
                    {channelDrillDown.type === 'appointments' && (
                      <>
                        <th className="p-2">Tipo</th>
                        <th className="p-2">Esito</th>
                        <th className="p-2">Store</th>
                      </>
                    )}
                    {channelDrillDown.type === 'trials' && (
                      <>
                        <th className="p-2">Audioprotesista</th>
                        <th className="p-2">Stato</th>
                      </>
                    )}
                    {channelDrillDown.type === 'sales' && (
                      <>
                        <th className="p-2">Audioprotesista</th>
                        <th className="p-2">Intermediario</th>
                        <th className="p-2 text-right">Valore</th>
                      </>
                    )}
                    {channelDrillDown.type === 'availability' && <th className="p-2">Store</th>}
                    <th className="p-2">Note</th>
                  </tr>
                </thead>
                <tbody>
                  {channelDrillDown.data.map((item, i) => (
                    <tr key={i} className="border-b border-gray-100 hover:bg-gray-50 transition-colors">
                      <td className="p-2 whitespace-nowrap">{item.data || item.fullDate || 'N/D'}</td>
                      <td className="p-2 font-bold">{item.nome || item.cliente || 'N/D'}</td>
                      <td className="p-2">{item.canale || 'N/D'}</td>
                      {channelDrillDown.type === 'patients' && <td className="p-2">{item.store || 'N/D'}</td>}
                      {channelDrillDown.type === 'calls' && <td className="p-2">{item.contatto || 'N/D'}</td>}
                      {channelDrillDown.type === 'appointments' && (
                        <>
                          <td className="p-2">{item.tipo || 'N/D'}</td>
                          <td className="p-2">{item.contatto || 'N/D'}</td>
                          <td className="p-2">{item.sede || 'N/D'}</td>
                        </>
                      )}
                      {channelDrillDown.type === 'trials' && (
                        <>
                          <td className="p-2">{item.audioprotesista || 'N/D'}</td>
                          <td className="p-2">{item.statoDocumento || item.contatto || 'N/D'}</td>
                        </>
                      )}
                      {channelDrillDown.type === 'sales' && (
                        <>
                          <td className="p-2">{item.audioprotesista || 'N/D'}</td>
                          <td className="p-2">{item.intermediario || 'N/D'}</td>
                          <td className="p-2 text-right font-bold">{(item.valore || 0).toLocaleString('it-IT')}</td>
                        </>
                      )}
                      {channelDrillDown.type === 'availability' && <td className="p-2">{item.storeName || 'N/D'}</td>}
                      <td className="p-2 opacity-60 truncate max-w-[200px]" title={item.note}>{item.note || '-'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
      {/* Drill-down Modal */}
      {drillDown && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-[200] flex items-center justify-center p-4">
          <div className="bg-white border border-[#141414] w-full max-w-6xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-gray-50">
              <div className="flex items-center gap-8">
                <div>
                  <h3 className="font-serif italic text-2xl">{drillDown.name}</h3>
                  <p className="text-[10px] uppercase tracking-widest opacity-50">
                    {drillDown.channel && <span className="text-emerald-600 font-bold">{drillDown.channel} • </span>}
                    {drillDown.month === -1 ? 'Intero Anno' : MONTHS[drillDown.month]} {drillDown.year || selectedYear}
                  </p>
                </div>
                
                <div className="flex gap-4 border-l border-[#141414]/10 pl-8">
                  <div className="text-center">
                    <p className="text-[9px] uppercase tracking-widest opacity-50 mb-1">Tot. Vendite</p>
                    <p className="font-bold text-emerald-600 text-lg">{drillDown.sales.reduce((acc: number, s: any) => acc + (s.valore || 0), 0).toLocaleString('it-IT')}</p>
                  </div>
                  <div className="text-center border-l border-[#141414]/10 pl-4">
                    <p className="text-[9px] uppercase tracking-widest opacity-50 mb-1">Tot. Storni</p>
                    <p className="font-bold text-red-500 text-lg">{drillDown.credits.reduce((acc: number, c: any) => acc + (c.valore || 0), 0).toLocaleString('it-IT')}</p>
                  </div>
                  <div className="text-center border-l border-[#141414]/10 pl-4">
                    <p className="text-[9px] uppercase tracking-widest opacity-50 mb-1">Risultato Netto</p>
                    <p className="font-bold text-[#141414] text-lg">{(drillDown.sales.reduce((acc: number, s: any) => acc + (s.valore || 0), 0) - drillDown.credits.reduce((acc: number, c: any) => acc + (c.valore || 0), 0)).toLocaleString('it-IT')}</p>
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-4">
                <button
                  onClick={handleDownloadDrillDownExcel}
                  className="flex items-center gap-2 px-4 py-2 border border-[#141414] hover:bg-black hover:text-white transition-all text-[10px] uppercase tracking-widest font-bold"
                >
                  <Download size={14} /> Download Excel
                </button>
                <button 
                  onClick={() => setDrillDown(null)}
                  className="w-10 h-10 flex items-center justify-center border border-[#141414] hover:bg-black hover:text-white transition-all"
                >
                  ✕
                </button>
              </div>
            </div>
            
            <div className="flex-1 overflow-y-auto p-6 space-y-8">
              {/* Sales Section */}
              <div>
                <h4 className="text-[10px] uppercase tracking-[0.2em] mb-4 flex items-center gap-2 text-emerald-600">
                  <TrendingUp size={14} /> Dettaglio Vendite ({drillDown.sales.length})
                </h4>
                {drillDown.sales.length > 0 ? (
                  <table className="w-full text-left border-collapse font-mono text-[11px]">
                    <thead>
                      <tr className="border-b border-[#141414] opacity-50">
                        <th className="py-2">Data</th>
                        <th className="py-2">Audioprotesista</th>
                        <th className="py-2">Cliente finale</th>
                        <th className="py-2">Canale</th>
                        <th className="py-2">Intermediario</th>
                        <th className="py-2 text-center">Quantità</th>
                        <th className="py-2 text-right">Importo</th>
                      </tr>
                    </thead>
                    <tbody>
                      {drillDown.sales.map((s, i) => (
                        <tr key={i} className="border-b border-gray-100 last:border-0 hover:bg-gray-50 transition-colors">
                          <td className="py-2">{s.fullDate}</td>
                          <td className="py-2 font-bold">{s.audioprotesista || 'N/D'}</td>
                          <td className="py-2">{s.cliente}</td>
                          <td className="py-2">{s.canale || 'N/D'}</td>
                          <td className="py-2">{s.intermediario || 'N/D'}</td>
                          <td className="py-2 text-center font-bold">{s.quantita || 0}</td>
                          <td className="py-2 text-right font-bold text-emerald-600">{s.valore.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <p className="text-xs italic opacity-40">Nessuna vendita registrata.</p>
                )}
              </div>

              {/* Credits Section */}
              <div>
                <h4 className="text-[10px] uppercase tracking-[0.2em] mb-4 flex items-center gap-2 text-red-500">
                  <AlertCircle size={14} /> Dettaglio Storni ({drillDown.credits.length})
                </h4>
                {drillDown.credits.length > 0 ? (
                  <table className="w-full text-left border-collapse font-mono text-[11px]">
                    <thead>
                      <tr className="border-b border-[#141414] opacity-50">
                        <th className="py-2">Data</th>
                        <th className="py-2">Audioprotesista</th>
                        <th className="py-2">Cliente finale</th>
                        <th className="py-2 text-center">Quantità</th>
                        <th className="py-2 text-right">Importo</th>
                      </tr>
                    </thead>
                    <tbody>
                      {drillDown.credits.map((c, i) => (
                        <tr key={i} className="border-b border-gray-100 last:border-0 hover:bg-red-50/30 transition-colors">
                          <td className="py-2">{c.fullDate}</td>
                          <td className="py-2 font-bold">{c.audioprotesista || 'N/D'}</td>
                          <td className="py-2">{c.cliente}</td>
                          <td className="py-2 text-center font-bold">{c.quantita || 0}</td>
                          <td className="py-2 text-right font-bold text-red-500">- {c.valore.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <p className="text-xs italic opacity-40">Nessuno storno registrato.</p>
                )}
              </div>
            </div>
            
            <div className="p-6 border-t border-[#141414] bg-gray-50">
              <button 
                onClick={() => setDrillDown(null)}
                className="w-full bg-[#141414] text-white py-3 font-bold hover:bg-[#333] transition-colors"
              >
                Chiudi Dettaglio
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Marketing Analysis Modal */}
      {showMarketingAnalysis && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white border border-[#141414] w-full max-w-6xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-[#141414] text-white">
              <div>
                <h3 className="font-serif italic text-2xl">Analisi Canali Marketing</h3>
                <p className="text-[10px] uppercase tracking-widest opacity-60">Suddivisione Vendite per Canale Primario • {selectedYear}</p>
              </div>
              <div className="flex items-center gap-4">
                <button
                  onClick={handleDownloadMarketingExcel}
                  className="flex items-center gap-2 px-4 py-2 border border-white/20 hover:bg-white hover:text-black transition-all text-[10px] uppercase tracking-widest font-bold"
                >
                  <Download size={14} /> Download Excel
                </button>
                <button 
                  onClick={() => setShowMarketingAnalysis(false)}
                  className="w-10 h-10 flex items-center justify-center border border-white/20 hover:bg-white hover:text-black transition-all"
                >
                  ✕
                </button>
              </div>
            </div>
            
            <div className="flex-1 overflow-auto p-6">
              <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead>
                  <tr className="bg-gray-50 border-b border-[#141414]">
                    <th className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 border-r border-[#141414] sticky left-0 bg-gray-50 z-10">Canale Primario</th>
                    <th className="p-4 text-[10px] uppercase tracking-widest font-bold opacity-80 text-center bg-gray-100 border-r border-[#141414]">Totale Anno</th>
                    {MONTHS.map(m => (
                      <th key={m} className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 text-center min-w-[100px]">
                        {m.slice(0, 3)}
                      </th>
                    ))}
                    <th className="p-4 text-[10px] uppercase tracking-widest font-bold opacity-80 text-center bg-gray-100 border-l border-[#141414]">TOTALE</th>
                  </tr>
                </thead>
                <tbody>
                  {(Array.from(new Set(marketingAnalysis.flatMap(m => m.channels.map((c: any) => c.name)))) as string[])
                    .sort()
                    .map(channelName => {
                      const channelTotal = marketingAnalysis.reduce((acc, m) => {
                        const c = m.channels.find((ch: any) => ch.name === channelName);
                        return acc + (c ? c.total : 0);
                      }, 0);

                      return (
                        <tr key={channelName} className="border-b border-gray-100 hover:bg-gray-50 transition-colors group">
                          <td className="p-4 text-xs font-bold border-r border-[#141414] sticky left-0 bg-white group-hover:bg-gray-50 z-10">
                            {channelName}
                          </td>
                          <td 
                            onClick={() => handleCellClick(null, -1, channelName)}
                            className="p-4 text-center bg-gray-50 font-bold border-r border-[#141414] cursor-pointer hover:bg-gray-100 transition-colors"
                          >
                            <div className="text-xs font-mono">
                              {channelTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                            </div>
                          </td>
                          {MONTHS.map((_, mIdx) => {
                            const monthData = marketingAnalysis.find(m => m.month === MONTHS[mIdx]);
                            const channelData = monthData?.channels.find((c: any) => c.name === channelName);
                            return (
                              <td 
                                key={mIdx} 
                                onClick={() => channelData && handleCellClick(null, mIdx, channelName)}
                                className={cn(
                                  "p-4 text-center transition-colors",
                                  channelData ? "cursor-pointer hover:bg-gray-50" : ""
                                )}
                              >
                                {channelData ? (
                                  <div className="space-y-1">
                                    <div className="text-xs font-mono font-bold">
                                      {channelData.total.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                                    </div>
                                    <div className="text-[9px] opacity-40 font-mono">
                                      {channelData.count} vend.
                                    </div>
                                  </div>
                                ) : (
                                  <span className="text-[10px] opacity-20">-</span>
                                )}
                              </td>
                            );
                          })}
                          <td 
                            onClick={() => handleCellClick(null, -1, channelName)}
                            className="p-4 text-center bg-gray-50 font-bold border-l border-[#141414] cursor-pointer hover:bg-gray-100 transition-colors"
                          >
                            <div className="text-xs font-mono">
                              € {channelTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                            </div>
                            <div className="text-[9px] opacity-40 font-mono font-normal">
                              {Math.round((channelTotal / marketingAnalysis.reduce((acc, m) => acc + m.channels.reduce((cAcc: number, c: any) => cAcc + c.total, 0), 0)) * 100)}%
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
                <tfoot>
                  <tr className="bg-[#141414] text-white font-bold">
                    <td className="p-4 text-[10px] uppercase tracking-widest sticky left-0 bg-[#141414] z-10">TOTALE MESE</td>
                    <td 
                      onClick={() => handleCellClick(null, -1)}
                      className="p-4 text-center text-xs font-mono bg-white/10 cursor-pointer hover:bg-white/20 transition-colors border-r border-white/20"
                    >
                      {marketingAnalysis.reduce((acc, m) => acc + m.channels.reduce((cAcc: number, c: any) => cAcc + c.total, 0), 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                    </td>
                    {MONTHS.map((_, mIdx) => {
                      const monthTotal = marketingAnalysis.find(m => m.month === MONTHS[mIdx])?.channels.reduce((acc: number, c: any) => acc + c.total, 0) || 0;
                      return (
                        <td 
                          key={mIdx} 
                          onClick={() => monthTotal !== 0 && handleCellClick(null, mIdx)}
                          className={cn(
                            "p-4 text-center text-xs font-mono transition-colors",
                            monthTotal !== 0 ? "cursor-pointer hover:bg-white/10" : ""
                          )}
                        >
                          {monthTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                        </td>
                      );
                    })}
                    <td 
                      onClick={() => handleCellClick(null, -1)}
                      className="p-4 text-center text-xs font-mono border-l border-white/20 cursor-pointer hover:bg-white/10 transition-colors"
                    >
                      {marketingAnalysis.reduce((acc, m) => acc + m.channels.reduce((cAcc: number, c: any) => cAcc + c.total, 0), 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
            
            <div className="p-6 border-t border-[#141414] bg-gray-50 flex justify-between items-center">
              <div className="flex items-center gap-6">
                <p className="text-[10px] uppercase tracking-widest opacity-50 italic">
                  * I dati includono solo vendite ≥ 600€ (ff.xlsx)
                </p>
                <button
                  onClick={handleDownloadMarketingExcel}
                  className="flex items-center gap-2 px-6 py-3 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold hover:bg-[#333] transition-all"
                >
                  <Download size={14} /> Download Excel
                </button>
              </div>
              <button 
                onClick={() => setShowMarketingAnalysis(false)}
                className="border border-[#141414] text-[#141414] px-8 py-3 font-bold hover:bg-gray-100 transition-colors"
              >
                Chiudi Analisi
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Screening Analysis Modal */}
      {showScreeningAnalysis && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white border border-[#141414] w-full max-w-6xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-[#141414] text-white">
              <div>
                <h3 className="font-serif italic text-2xl">Analisi Screening</h3>
                <p className="text-[10px] uppercase tracking-widest opacity-60">Dettaglio Depositi e Canali Secondari • {selectedYear}</p>
              </div>
              <div className="flex items-center gap-4">
                <button
                  onClick={handleDownloadScreeningExcel}
                  className="flex items-center gap-2 px-4 py-2 border border-white/20 hover:bg-white hover:text-black transition-all text-[10px] uppercase tracking-widest font-bold"
                >
                  <Download size={14} /> Download Excel
                </button>
                <button 
                  onClick={() => setShowScreeningAnalysis(false)}
                  className="w-10 h-10 flex items-center justify-center border border-white/20 hover:bg-white hover:text-black transition-all"
                >
                  ✕
                </button>
              </div>
            </div>
            
            <div className="flex-1 overflow-auto p-6">
              <table className="w-full text-left border-collapse min-w-[1000px]">
                <thead>
                  <tr className="bg-gray-50 border-b border-[#141414]">
                    <th className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 border-r border-[#141414] sticky left-0 bg-gray-50 z-10">Deposito di origine / Canale Secondario</th>
                    <th className="p-4 text-[10px] uppercase tracking-widest font-bold opacity-80 text-center bg-gray-100 border-r border-[#141414]">Totale Anno</th>
                    {MONTHS.map(m => (
                      <th key={m} className="p-4 text-[10px] uppercase tracking-widest font-normal opacity-50 text-center min-w-[100px]">
                        {m.slice(0, 3)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="font-mono text-[11px]">
                  {screeningAnalysis.map((dep, depIdx) => (
                    <React.Fragment key={depIdx}>
                      <tr className="border-b border-[#141414] bg-gray-50/50">
                        <td className="p-4 font-bold border-r border-[#141414] sticky left-0 bg-gray-50 z-10">
                          {dep.name}
                        </td>
                        <td 
                          onClick={() => handleScreeningCellClick(dep.name, -1)}
                          className="p-4 text-center font-bold border-r border-[#141414] bg-gray-100/50 cursor-pointer hover:bg-gray-200 transition-colors"
                        >
                          {dep.total.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                        </td>
                        {MONTHS.map((_, mIdx) => {
                          const monthTotal = dep.channels.reduce((acc: number, c: any) => acc + c.monthly[mIdx], 0);
                          return (
                            <td 
                              key={mIdx} 
                              onClick={() => monthTotal !== 0 && handleScreeningCellClick(dep.name, mIdx)}
                              className={cn(
                                "p-4 text-center font-bold transition-colors",
                                monthTotal !== 0 ? "cursor-pointer hover:bg-gray-100" : ""
                              )}
                            >
                              {monthTotal !== 0 ? monthTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 }) : '-'}
                            </td>
                          );
                        })}
                      </tr>
                      {dep.channels.map((c: any, cIdx: number) => (
                        <tr key={cIdx} className="border-b border-gray-100 hover:bg-gray-50 transition-colors group">
                          <td className="p-4 pl-8 text-xs border-r border-[#141414] sticky left-0 bg-white group-hover:bg-gray-50 z-10 italic opacity-80">
                            ↳ {c.name}
                          </td>
                          <td 
                            onClick={() => handleScreeningCellClick(dep.name, -1, c.name)}
                            className="p-4 text-center border-r border-[#141414] opacity-80 cursor-pointer hover:bg-gray-100 transition-colors"
                          >
                            € {c.total.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                          </td>
                          {MONTHS.map((_, mIdx) => (
                            <td 
                              key={mIdx} 
                              onClick={() => c.monthly[mIdx] !== 0 && handleScreeningCellClick(dep.name, mIdx, c.name)}
                              className={cn(
                                "p-4 text-center opacity-60 transition-colors",
                                c.monthly[mIdx] !== 0 ? "cursor-pointer hover:bg-gray-100" : ""
                              )}
                            >
                              {c.monthly[mIdx] !== 0 ? c.monthly[mIdx].toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 }) : '-'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </React.Fragment>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-[#141414] text-white font-bold">
                    <td className="p-4 text-[10px] uppercase tracking-widest sticky left-0 bg-[#141414] z-10">TOTALE GENERALE SCREENING</td>
                    <td 
                      onClick={() => handleScreeningCellClick(null, -1)}
                      className="p-4 text-center text-xs font-mono bg-white/10 border-r border-white/20 cursor-pointer hover:bg-white/20 transition-colors"
                    >
                      € {screeningAnalysis.reduce((acc, dep) => acc + dep.total, 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                    </td>
                    {MONTHS.map((_, mIdx) => {
                      const monthTotal = screeningAnalysis.reduce((acc, dep) => 
                        acc + dep.channels.reduce((cAcc: number, c: any) => cAcc + c.monthly[mIdx], 0), 0);
                      return (
                        <td 
                          key={mIdx} 
                          onClick={() => monthTotal !== 0 && handleScreeningCellClick(null, mIdx)}
                          className={cn(
                            "p-4 text-center text-xs font-mono transition-colors",
                            monthTotal !== 0 ? "cursor-pointer hover:bg-white/10" : ""
                          )}
                        >
                          € {monthTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                        </td>
                      );
                    })}
                  </tr>
                </tfoot>
              </table>
            </div>
            
            <div className="p-6 border-t border-[#141414] bg-gray-50 flex justify-between items-center">
              <p className="text-[10px] uppercase tracking-widest opacity-50 italic">
                * Analisi basata su Canale Primario: "Screening" o "Promoter"
              </p>
              <button 
                onClick={() => setShowScreeningAnalysis(false)}
                className="bg-[#141414] text-white px-8 py-3 font-bold hover:bg-[#333] transition-colors"
              >
                Chiudi Analisi
              </button>
            </div>
          </div>
        </div>
      )}

      {showMediciAnalysis && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white border border-[#141414] w-full max-w-[95vw] max-h-[95vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-gray-50">
              <div>
                <h3 className="font-serif italic text-3xl">Analisi Medici</h3>
                <p className="text-[10px] uppercase tracking-widest opacity-50">
                  Riepilogo Vendite per Audioprotesista e Canale Secondario • {selectedYear}
                </p>
              </div>
              <button 
                onClick={() => setShowMediciAnalysis(false)}
                className="w-12 h-12 flex items-center justify-center border border-[#141414] hover:bg-black hover:text-white transition-all text-xl"
              >
                ✕
              </button>
            </div>
            
            <div className="flex-1 overflow-auto">
              <table className="w-full text-left border-collapse font-mono text-[11px]">
                <thead className="sticky top-0 z-20 bg-[#141414] text-white">
                  <tr>
                    <th className="p-4 border-r border-white/20 sticky left-0 bg-[#141414] z-30 min-w-[250px]">Audioprotesista / Canale Secondario</th>
                    <th className="p-4 text-center border-r border-white/20 bg-white/10">Totale Anno</th>
                    {MONTHS.map((m, idx) => (
                      <th key={idx} className="p-4 text-center border-r border-white/10 last:border-r-0">{m}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {mediciAnalysis.map((dep, idx) => (
                    <React.Fragment key={idx}>
                      <tr className="border-b border-[#141414] bg-gray-50/50">
                        <td className="p-4 font-bold border-r border-[#141414] sticky left-0 bg-gray-50 z-10">
                          {dep.name}
                        </td>
                        <td 
                          onClick={() => handleMediciCellClick(dep.name, -1)}
                          className="p-4 text-center font-bold border-r border-[#141414] bg-gray-100/50 cursor-pointer hover:bg-gray-200 transition-colors"
                        >
                          {dep.total.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                        </td>
                        {MONTHS.map((_, mIdx) => {
                          const monthTotal = dep.channels.reduce((acc: number, c: any) => acc + c.monthly[mIdx], 0);
                          return (
                            <td 
                              key={mIdx} 
                              onClick={() => monthTotal !== 0 && handleMediciCellClick(dep.name, mIdx)}
                              className={cn(
                                "p-4 text-center font-bold transition-colors",
                                monthTotal !== 0 ? "cursor-pointer hover:bg-gray-100" : ""
                              )}
                            >
                              {monthTotal !== 0 ? monthTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 }) : '-'}
                            </td>
                          );
                        })}
                      </tr>
                      {dep.channels.map((c: any, cIdx: number) => (
                        <tr key={cIdx} className="border-b border-gray-100 hover:bg-gray-50 transition-colors group">
                          <td className="p-4 pl-8 text-xs border-r border-[#141414] sticky left-0 bg-white group-hover:bg-gray-50 z-10 italic opacity-80">
                            ↳ {c.name}
                          </td>
                          <td 
                            onClick={() => handleMediciCellClick(dep.name, -1, c.name)}
                            className="p-4 text-center border-r border-[#141414] opacity-80 cursor-pointer hover:bg-gray-100 transition-colors"
                          >
                            € {c.total.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                          </td>
                          {MONTHS.map((_, mIdx) => (
                            <td 
                              key={mIdx} 
                              onClick={() => c.monthly[mIdx] !== 0 && handleMediciCellClick(dep.name, mIdx, c.name)}
                              className={cn(
                                "p-4 text-center opacity-60 transition-colors",
                                c.monthly[mIdx] !== 0 ? "cursor-pointer hover:bg-gray-100" : ""
                              )}
                            >
                              {c.monthly[mIdx] !== 0 ? `€ ${c.monthly[mIdx].toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}` : '-'}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </React.Fragment>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-[#141414] text-white font-bold">
                    <td className="p-4 text-[10px] uppercase tracking-widest sticky left-0 bg-[#141414] z-10">TOTALE GENERALE MEDICI</td>
                    <td 
                      onClick={() => handleMediciCellClick(null, -1)}
                      className="p-4 text-center text-xs font-mono bg-white/10 border-r border-white/20 cursor-pointer hover:bg-white/20 transition-colors"
                    >
                      € {mediciAnalysis.reduce((acc, dep) => acc + dep.total, 0).toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                    </td>
                    {MONTHS.map((_, mIdx) => {
                      const monthTotal = mediciAnalysis.reduce((acc, dep) => 
                        acc + dep.channels.reduce((cAcc: number, c: any) => cAcc + c.monthly[mIdx], 0), 0);
                      return (
                        <td 
                          key={mIdx} 
                          onClick={() => monthTotal !== 0 && handleMediciCellClick(null, mIdx)}
                          className={cn(
                            "p-4 text-center text-xs font-mono transition-colors",
                            monthTotal !== 0 ? "cursor-pointer hover:bg-white/10" : ""
                          )}
                        >
                          € {monthTotal.toLocaleString('it-IT', { useGrouping: true, maximumFractionDigits: 0 })}
                        </td>
                      );
                    })}
                  </tr>
                </tfoot>
              </table>
            </div>
            
            <div className="p-6 border-t border-[#141414] bg-gray-50 flex justify-between items-center">
              <div className="flex items-center gap-6">
                <p className="text-[10px] uppercase tracking-widest opacity-50 italic">
                  * Analisi basata su Canale Primario: "Medico"
                </p>
                <button
                  onClick={handleDownloadMediciExcel}
                  className="flex items-center gap-2 px-6 py-3 bg-[#141414] text-white text-[10px] uppercase tracking-widest font-bold hover:bg-[#333] transition-all"
                >
                  <Download size={14} /> Download Excel
                </button>
              </div>
              <button 
                onClick={() => setShowMediciAnalysis(false)}
                className="bg-[#141414] text-white px-8 py-3 font-bold hover:bg-[#333] transition-colors"
              >
                Chiudi Analisi
              </button>
            </div>
          </div>
        </div>
      )}
      {isBudgetSummaryOpen && (
        <div className="fixed inset-0 bg-[#141414]/90 z-[1000] p-4 lg:p-8 flex items-center justify-center overflow-auto">
          <div className="bg-white w-full max-w-[1600px] border border-[#141414] shadow-2xl flex flex-col max-h-full">
            <div className="p-4 border-b border-[#141414] flex justify-between items-center bg-gray-50 shrink-0">
              <div className="flex items-center gap-6">
                <h2 className="font-serif italic text-xl flex items-center gap-2">
                  <FileSpreadsheet size={20} /> Situazione Budget {selectedYear}
                </h2>
                <button
                  onClick={handleDownloadBudgetSituation}
                  className="bg-[#141414] text-white px-4 py-2 text-[10px] uppercase tracking-widest font-bold hover:bg-[#333] transition-all flex items-center gap-2"
                >
                  <Download size={14} /> Download Excel
                </button>
              </div>
              <button 
                onClick={() => setIsBudgetSummaryOpen(false)}
                className="p-2 hover:bg-gray-200 transition-colors uppercase tracking-widest text-[10px] font-bold flex items-center gap-2"
              >
                Chiudi <X size={16} />
              </button>
            </div>

            <div className="p-6 overflow-y-auto space-y-12">
              {[1, 2, 3, 4].map(quarter => {
                const monthsStart = (quarter - 1) * 3;
                const monthIndices = [monthsStart, monthsStart + 1, monthsStart + 2];
                const qLabels = ["I TRIMESTRE", "II TRIMESTRE", "III TRIMESTRE", "IV TRIMESTRE"];
                const qLabel = qLabels[quarter - 1];
                
                // Calculate quarterly totals for all personnel
                const qActualTotal = budgetSituationData.reduce((acc, p) => acc + monthIndices.reduce((mAcc, mIdx) => mAcc + p.months[mIdx].actual, 0), 0);
                const qBudgetTotal = budgetSituationData.reduce((acc, p) => acc + monthIndices.reduce((mAcc, mIdx) => mAcc + p.months[mIdx].budget, 0), 0);
                const qTotalDiff = qActualTotal - qBudgetTotal;
                const qTotalPercValue = qBudgetTotal > 0 ? (qActualTotal / qBudgetTotal - 1) * 100 : 0;

                return (
                  <div key={quarter} className="overflow-x-auto">
                    <table className="w-full border-collapse border border-[#141414] text-[10px] font-mono whitespace-nowrap">
                      <thead>
                        {/* Quarter Header Row */}
                        <tr>
                          <th className="border border-[#141414] p-1 bg-white"></th>
                          {monthIndices.map(mIdx => (
                            <th key={mIdx} colSpan={3} className="border border-[#141414] p-1 bg-yellow-400 uppercase tracking-widest">
                              {MONTHS[mIdx]}
                            </th>
                          ))}
                          <th colSpan={3} className="border border-[#141414] p-1 bg-yellow-400 uppercase tracking-widest italic font-serif">
                            {qLabel}
                          </th>
                          <th className="border border-[#141414] p-1 bg-white"></th>
                        </tr>
                        {/* Sub-headers Row */}
                        <tr className="bg-blue-50">
                          <th className="border border-[#141414] p-2 text-left bg-white font-serif italic text-sm min-w-[200px]">Audioprotesista</th>
                          {monthIndices.map(mIdx => (
                            <React.Fragment key={mIdx}>
                              <th className="border border-[#141414] p-1 text-center bg-blue-100/50">fatturato</th>
                              <th className="border border-[#141414] p-1 text-center bg-blue-100/50">bgt</th>
                              <th className="border border-[#141414] p-1 text-center bg-blue-100/50">fatt/ bgt</th>
                            </React.Fragment>
                          ))}
                          <th className="border border-[#141414] p-1 text-center bg-blue-100/50">fatturato</th>
                          <th className="border border-[#141414] p-1 text-center bg-blue-100/50">bgt</th>
                          <th className="border border-[#141414] p-1 text-center bg-blue-100/50">fatt/ bgt</th>
                          <th className="border border-[#141414] p-1 text-center bg-white">fatt / bgt</th>
                        </tr>
                      </thead>
                      <tbody>
                        {budgetSituationData.map(person => {
                          const pQActual = monthIndices.reduce((acc, mIdx) => acc + person.months[mIdx].actual, 0);
                          const pQBudget = monthIndices.reduce((acc, mIdx) => acc + person.months[mIdx].budget, 0);
                          const pQDiff = pQActual - pQBudget;
                          const pQPercValue = pQBudget > 0 ? (pQActual / pQBudget - 1) * 100 : 0;

                          return (
                            <tr key={person.name} className="hover:bg-gray-50 transition-colors">
                              <td className="border border-[#141414] p-2 font-bold font-serif text-xs">{person.name}</td>
                              {monthIndices.map(mIdx => {
                                const m = person.months[mIdx];
                                return (
                                  <React.Fragment key={mIdx}>
                                    <td className="border border-[#141414] p-1 text-right">{Math.round(m.actual).toLocaleString('it-IT')}</td>
                                    <td className="border border-[#141414] p-1 text-right">{Math.round(m.budget).toLocaleString('it-IT')}</td>
                                    <td className={cn(
                                      "border border-[#141414] p-1 text-right font-bold",
                                      m.diff >= 0 ? "text-emerald-700" : "text-red-700"
                                    )}>
                                      {Math.round(m.diff).toLocaleString('it-IT')}
                                    </td>
                                  </React.Fragment>
                                );
                              })}
                              {/* Quarterly Columns */}
                              <td className="border border-[#141414] p-1 text-right bg-gray-50">{Math.round(pQActual).toLocaleString('it-IT')}</td>
                              <td className="border border-[#141414] p-1 text-right bg-gray-50">{Math.round(pQBudget).toLocaleString('it-IT')}</td>
                              <td className={cn(
                                "border border-[#141414] p-1 text-right font-bold bg-gray-50",
                                pQDiff >= 0 ? "text-emerald-700" : "text-red-700"
                              )}>
                                {Math.round(pQDiff).toLocaleString('it-IT')}
                              </td>
                              {/* Quarterly Percentage */}
                              <td className={cn(
                                "border border-[#141414] p-1 text-center font-bold",
                                pQPercValue >= 0 ? "text-emerald-600" : "text-red-600"
                              )}>
                                {pQPercValue.toFixed(0)}%
                              </td>
                            </tr>
                          );
                        })}
                        {/* Totals Row */}
                        <tr className="bg-gray-100 font-bold border-t-2 border-[#141414]">
                          <td className="border border-[#141414] p-2 uppercase tracking-widest text-[9px]">TOTALI</td>
                          {monthIndices.map(mIdx => {
                            const mActual = budgetSituationData.reduce((acc, p) => acc + p.months[mIdx].actual, 0);
                            const mBudget = budgetSituationData.reduce((acc, p) => acc + p.months[mIdx].budget, 0);
                            return (
                              <React.Fragment key={mIdx}>
                                <td className="border border-[#141414] p-1 text-right">{Math.round(mActual).toLocaleString('it-IT')}</td>
                                <td className="border border-[#141414] p-1 text-right">{Math.round(mBudget).toLocaleString('it-IT')}</td>
                                <td className={cn(
                                  "border border-[#141414] p-1 text-right",
                                  (mActual - mBudget) >= 0 ? "text-emerald-700" : "text-red-700"
                                )}>
                                  {Math.round(mActual - mBudget).toLocaleString('it-IT')}
                                </td>
                              </React.Fragment>
                            );
                          })}
                          {/* Quarter Totals */}
                          <td className="border border-[#141414] p-1 text-right bg-gray-200">{Math.round(qActualTotal).toLocaleString('it-IT')}</td>
                          <td className="border border-[#141414] p-1 text-right bg-gray-200">{Math.round(qBudgetTotal).toLocaleString('it-IT')}</td>
                          <td className={cn(
                            "border border-[#141414] p-1 text-right bg-gray-200",
                            qTotalDiff >= 0 ? "text-emerald-700" : "text-red-700"
                          )}>
                            {Math.round(qTotalDiff).toLocaleString('it-IT')}
                          </td>
                          <td className={cn(
                            "border border-[#141414] p-1 text-center font-bold bg-gray-200",
                            qTotalPercValue >= 0 ? "text-emerald-600" : "text-red-600"
                          )}>
                            {qTotalPercValue.toFixed(0)}%
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
            
            <div className="p-4 border-t border-[#141414] bg-gray-50 shrink-0 text-right">
              <p className="text-[10px] opacity-40 italic uppercase tracking-wider">
                * Dati calcolati su Vendite (ff.xlsx) e Storni (rr.xlsx) caricati nel sistema.
              </p>
            </div>
          </div>
        </div>
      )}

      {showCoefficientiAnalysis && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white border border-[#141414] w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl">
            <div className="p-6 border-b border-[#141414] flex justify-between items-center bg-gray-50">
              <div>
                <h3 className="font-serif italic text-3xl">Coefficienti di Vendita</h3>
                <p className="text-[10px] uppercase tracking-widest opacity-50">
                  Media Vendita • Apparecchi • Coefficiente Binaurale • {selectedYear}
                </p>
              </div>
              <button 
                onClick={() => setShowCoefficientiAnalysis(false)}
                className="w-12 h-12 flex items-center justify-center border border-[#141414] hover:bg-black hover:text-white transition-all text-xl"
              >
                ✕
              </button>
            </div>
            
            <div className="flex-1 overflow-auto p-6">
              <table className="w-full text-left border-collapse font-mono text-xs">
                <thead className="sticky top-0 bg-[#141414] text-white">
                  <tr>
                    <th className="p-4 border-r border-white/20">Audioprotesista</th>
                    <th className="p-4 text-right border-r border-white/20">Fatturato Netto</th>
                    <th className="p-4 text-center border-r border-white/20">Apparecchi (PZ)</th>
                    <th className="p-4 text-right border-r border-white/20">Media Vendita</th>
                    <th className="p-4 text-center border-r border-white/20">Clienti</th>
                    <th className="p-4 text-center bg-emerald-600">Coeff. Binaurale</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {coefficientiAnalysis.map((row, idx) => (
                    <tr key={idx} className="hover:bg-gray-50 transition-colors">
                      <td className="p-4 font-bold border-r border-gray-100 uppercase tracking-tighter">{row.name}</td>
                      <td 
                        className="p-4 text-right border-r border-gray-100 cursor-pointer hover:bg-gray-100 transition-colors"
                        onClick={() => handleCellClick(row.name, -1, undefined, selectedYear, 'audioprotesista')}
                      >
                        € {row.netValore.toLocaleString('it-IT', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}
                      </td>
                      <td 
                        className="p-4 text-center border-r border-gray-100 font-bold cursor-pointer hover:bg-gray-100 transition-colors"
                        onClick={() => handleCellClick(row.name, -1, undefined, selectedYear, 'audioprotesista')}
                      >
                        {row.totalQuantita}
                      </td>
                      <td 
                        className="p-4 text-right border-r border-gray-100 font-bold cursor-pointer hover:bg-gray-100 transition-colors"
                        onClick={() => handleCellClick(row.name, -1, undefined, selectedYear, 'audioprotesista')}
                      >
                        € {row.mediaVendita.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </td>
                      <td 
                        className="p-4 text-center border-r border-gray-100 opacity-60 cursor-pointer hover:bg-gray-100 transition-colors"
                        onClick={() => handleCellClick(row.name, -1, undefined, selectedYear, 'audioprotesista')}
                      >
                        {row.numCustomers}
                      </td>
                      <td 
                        className="p-4 text-center font-bold text-sm bg-emerald-50 text-emerald-700 cursor-pointer hover:bg-emerald-100 transition-colors"
                        onClick={() => handleCellClick(row.name, -1, undefined, selectedYear, 'audioprotesista')}
                      >
                        {row.coefficienteBinaurale.toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot className="sticky bottom-0 bg-gray-100 font-bold text-sm">
                  <tr>
                    <td className="p-4 uppercase tracking-widest text-[10px]">Totale Generale</td>
                    <td 
                      className="p-4 text-right cursor-pointer hover:bg-gray-200 transition-colors"
                      onClick={() => handleCellClick(null, -1, undefined, selectedYear, 'audioprotesista')}
                    >
                      € {coefficientiAnalysis.reduce((acc, r) => acc + r.netValore, 0).toLocaleString('it-IT', { maximumFractionDigits: 0 })}
                    </td>
                    <td 
                      className="p-4 text-center cursor-pointer hover:bg-gray-200 transition-colors"
                      onClick={() => handleCellClick(null, -1, undefined, selectedYear, 'audioprotesista')}
                    >
                      {coefficientiAnalysis.reduce((acc, r) => acc + r.totalQuantita, 0)}
                    </td>
                    <td 
                      className="p-4 text-right cursor-pointer hover:bg-gray-200 transition-colors"
                      onClick={() => handleCellClick(null, -1, undefined, selectedYear, 'audioprotesista')}
                    >
                      € {(coefficientiAnalysis.reduce((acc, r) => acc + r.netValore, 0) / (coefficientiAnalysis.reduce((acc, r) => acc + r.totalQuantita, 0) || 1)).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </td>
                    <td 
                      className="p-4 text-center cursor-pointer hover:bg-gray-200 transition-colors"
                      onClick={() => handleCellClick(null, -1, undefined, selectedYear, 'audioprotesista')}
                    >
                      {coefficientiAnalysis.reduce((acc, r) => acc + r.numCustomers, 0)}
                    </td>
                    <td 
                      className="p-4 text-center bg-emerald-100 text-emerald-800 cursor-pointer hover:bg-emerald-200 transition-colors"
                      onClick={() => handleCellClick(null, -1, undefined, selectedYear, 'audioprotesista')}
                    >
                      {(coefficientiAnalysis.reduce((acc, r) => acc + r.totalQuantita, 0) / (coefficientiAnalysis.reduce((acc, r) => acc + r.numCustomers, 0) || 1)).toLocaleString('it-IT', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                    </td>
                  </tr>
                </tfoot>
              </table>
              <div className="mt-8 p-4 bg-blue-50 border-l-4 border-blue-500 text-[10px] space-y-1">
                <p className="font-bold uppercase tracking-widest text-blue-800 mb-1">Legenda Calcoli:</p>
                <p>• <span className="font-bold">Apparecchi:</span> Somma "Quantità totale" (min: 1, max: 2 per riga di vendita).</p>
                <p>• <span className="font-bold">Media Vendita:</span> Fatturato Netto (Vendite - Storni) / Totale Apparecchi.</p>
                <p>• <span className="font-bold">Coeff. Binaurale:</span> Totale Apparecchi / Numero Clienti Unici.</p>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
