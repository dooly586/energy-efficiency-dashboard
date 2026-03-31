import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  LineChart, Line, ComposedChart
} from 'recharts';
import { 
  Upload, Download, Calculator, Zap, DollarSign, TrendingDown, Calendar, Settings, FileSpreadsheet
} from 'lucide-react';
const formatMoney = (val) => new Intl.NumberFormat('ko-KR').format(Math.round(val || 0));

export default function App() {
  const [inputMode, setInputMode] = useState('excel'); // 'excel' or 'manual'
  const [rawData, setRawData] = useState([]);
  const [fileName, setFileName] = useState('');
  const [years, setYears] = useState([]);
  const [startYear, setStartYear] = useState('');
  const [startMonth, setStartMonth] = useState('');

  // Manual Input Mode: 'monthly' or 'annual'
  const [manualInputType, setManualInputType] = useState('annual');

  // Annual Quick Input
  const [annualBill, setAnnualBill] = useState('');

  // Manual Input Data (12 months)
  const [manualData, setManualData] = useState(
    Array.from({ length: 12 }, (_, i) => ({
      month: i + 1,
      usage: '',
      bill: ''
    }))
  );

  // Simulation Parameters
  const [calculationMethod, setCalculationMethod] = useState('billing'); // 'billing' or 'theoretical'
  const [otherPowerBill, setOtherPowerBill] = useState(''); // 월 비보일러 전기요금

  // 이론 계산 방식용 파라미터
  const [boilerCapacity, setBoilerCapacity] = useState(''); // 전기용량 (kW)
  const [dailyUsageHours, setDailyUsageHours] = useState(''); // 일평균 사용시간 (시간/일)
  const [annualOperatingDays, setAnnualOperatingDays] = useState('365'); // 년간 가동일 (일/년)

  // 전기요금 계약종별 (2026년 한전 요금표 기준)
  const [contractType, setContractType] = useState('general'); // 계약종별
  const [customRate, setCustomRate] = useState(''); // 직접입력 요금
  const [rateSource, setRateSource] = useState('actual'); // 'actual' (실제 데이터) or 'contract' (계약종별)

  // 상세 시설 투자비
  const [boilerCount, setBoilerCount] = useState(1);
  const [boilerUnitPrice, setBoilerUnitPrice] = useState('');
  const [installationCost, setInstallationCost] = useState('');
  const [thermalTankCost, setThermalTankCost] = useState(''); // 축열탱크비
  const [electricalWorkCost, setElectricalWorkCost] = useState('');
  const [otherCosts, setOtherCosts] = useState('');

  const [kepcoRate, setKepcoRate] = useState(''); // 켑코이에스 비율
  const [sgiRate, setSgiRate] = useState(''); // SGI 비율

  const [installmentMonths, setInstallmentMonths] = useState(60); // 할부개월수
  const [savingsRate, setSavingsRate] = useState(30); // 예상 절감율 (%)

  const [googleSheetUrl, setGoogleSheetUrl] = useState('');
  const [isLoadingSheet, setIsLoadingSheet] = useState(false);

  const fileInputRef = useRef(null);
  const infographicRef = useRef(null);

  // 한전 전기요금 계약종별 (2026년 4월 16일 시행 기준 평균 단가)
  const contractTypes = {
    residential: { name: '주택용 (저압)', rate: 150, description: '일반 가정, 누진제 적용' },
    general: { name: '일반용 (저압)', rate: 130, description: '상업시설, 사무실 등' },
    industrial: { name: '산업용 (고압)', rate: 95, description: '공장, 제조업 등' },
    agricultural_low: { name: '농사용 갑 (저압)', rate: 60, description: '저압 농업용' },
    agricultural_high: { name: '농사용 을 (고압)', rate: 55, description: '고압 농업용' },
    educational: { name: '교육용', rate: 85, description: '학교, 교육시설' },
    custom: { name: '직접입력', rate: 0, description: '사용자 지정 단가' }
  };

  // 실제 데이터에서 평균 전기단가 계산
  const getActualAverageRate = () => {
    if (rawData.length === 0) return null;
    const totalBill = rawData.reduce((sum, d) => sum + d.bill, 0);
    const totalUsage = rawData.reduce((sum, d) => sum + d.usage, 0);
    if (totalUsage === 0) return null;
    return Math.round(totalBill / totalUsage * 10) / 10; // 소수점 1자리
  };

  // 선택된 계약종별의 전기단가 계산
  const getElectricityRate = () => {
    const actualRate = getActualAverageRate();

    // Excel 데이터가 있고, 사용자가 실제 데이터 사용을 선택한 경우
    if (actualRate !== null && rateSource === 'actual') {
      return actualRate;
    }

    // 계약종별 단가 사용 (Excel 데이터가 없거나, 사용자가 계약종별 선택)
    if (contractType === 'custom') {
      return Number(customRate) || 120;
    }
    return contractTypes[contractType]?.rate || 120;
  };

  const handleMoneyChange = (setter) => (e) => setter(e.target.value.replace(/[^0-9]/g, ''));
  const formatInput = (val) => val ? Number(val).toLocaleString() : '';

  // Handle Manual Data Change
  const handleManualDataChange = (index, field, value) => {
    const cleanValue = value.replace(/[^0-9]/g, '');
    setManualData(prev => {
      const newData = [...prev];
      newData[index] = { ...newData[index], [field]: cleanValue };
      return newData;
    });
  };

  // Apply Annual Data (Quick Input)
  const applyAnnualData = () => {
    const currentYear = Number(startYear) || new Date().getFullYear();
    const totalBill = Number(annualBill) || 0;

    if (!totalBill) {
      alert('연간 전기 요금을 입력해주세요.');
      return;
    }

    // Estimate usage based on selected contract type rate
    const estimatedTotalUsage = Math.round(totalBill / getElectricityRate());

    // Divide by 12 months
    const monthlyUsage = Math.round(estimatedTotalUsage / 12);
    const monthlyBillAmount = Math.round(totalBill / 12);

    const parsed = Array.from({ length: 12 }, (_, i) => ({
      year: currentYear,
      month: i + 1,
      usage: monthlyUsage,
      bill: monthlyBillAmount
    }));

    setRawData(parsed);
    const uniqueYears = [...new Set(parsed.map(item => item.year))].sort((a,b)=>b-a);
    setYears(uniqueYears);
    if (!startMonth) {
      setStartMonth(1);
    }
  };

  // Apply Manual Data (Monthly Input)
  const applyManualData = () => {
    const currentYear = Number(startYear) || new Date().getFullYear();
    const parsed = manualData
      .filter(d => d.usage && d.bill)
      .map(d => ({
        year: currentYear,
        month: d.month,
        usage: Number(d.usage),
        bill: Number(d.bill)
      }));

    if (parsed.length === 0) {
      alert('최소 1개월 이상의 사용량과 요금을 입력해주세요.');
      return;
    }

    setRawData(parsed);
    const uniqueYears = [...new Set(parsed.map(item => item.year))].sort((a,b)=>b-a);
    setYears(uniqueYears);
    if (!startMonth) {
      setStartMonth(1);
    }
  };

  // 공통 시트 데이터 파싱 함수
  const parseSheetData = (data, sourceName = '') => {
    let parsed = [];

    // 처음 10줄 로그
    console.log(`📄 [${sourceName}] 데이터 처음 10줄:`);
    for (let i = 0; i < Math.min(10, data.length); i++) {
      console.log(`  ${i}:`, data[i]);
    }

    // ── 1. 헤더 행 찾기 (점수제: 3그룹 중 2그룹 이상 포함하면 헤더 인정) ──
    let headerRowIndex = -1;
    let bestScore = 0;

    for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      if (!row || !row.some(c => c != null && c !== '')) continue;
      const rowStr = row.map(c => String(c || '')).join(' ').toLowerCase();

      let score = 0;
      if (rowStr.includes('월') || rowStr.includes('년') || rowStr.includes('date') || rowStr.includes('날짜')) score += 2;
      if (rowStr.includes('사용') || rowStr.includes('kwh') || rowStr.includes('전력')) score += 2;
      if (rowStr.includes('요금') || rowStr.includes('금액') || rowStr.includes('청구')) score += 2;

      if (score > bestScore) {
        bestScore = score;
        headerRowIndex = i;
      }
    }

    if (bestScore < 2) headerRowIndex = -1;
    console.log(`헤더 행: ${headerRowIndex} (점수: ${bestScore})`);

    let monthIndex = -1, yearIndex = -1, usageIndex = -1, billIndex = -1;

    if (headerRowIndex !== -1) {
      const headerRow = data[headerRowIndex];
      console.log('헤더 내용:', headerRow);

      // Pass 1 – 우선순위 요금 컬럼 (청구/납부/합계/총)
      headerRow.forEach((col, idx) => {
        if (col == null || col === '') return;
        const cStr = String(col).toLowerCase().replace(/\s/g, '');
        if ((cStr.includes('청구') || cStr.includes('납부') || cStr.includes('합계') ||
             cStr.includes('최종') || cStr.includes('총액')) &&
            !cStr.includes('유형') && !cStr.includes('타입') && !cStr.includes('기간')) {
          billIndex = idx;
          console.log(`⭐ 요금(우선): ${idx} (${col})`);
        }
      });

      // Pass 2 – 나머지 컬럼
      headerRow.forEach((col, idx) => {
        if (col == null || col === '') return;
        const cStr = String(col).toLowerCase().replace(/\s/g, '');

        // 사용량
        if (usageIndex === -1 && !cStr.includes('기간') && !cStr.includes('계약') &&
            (cStr.includes('사용량') || cStr.includes('kwh') || cStr.includes('전력량'))) {
          usageIndex = idx;
          console.log(`사용량: ${idx} (${col})`);
        }
        // 요금 2순위
        if (billIndex === -1 &&
            (cStr.includes('요금') || cStr.includes('금액')) &&
            !cStr.includes('유형') && !cStr.includes('타입') &&
            !cStr.includes('기본') && !cStr.includes('전력량') && !cStr.includes('부가')) {
          billIndex = idx;
          console.log(`요금: ${idx} (${col})`);
        }
        // 년월 통합
        if (monthIndex === -1 && (cStr.includes('년월') || cStr.includes('날짜') || cStr === 'date')) {
          monthIndex = idx;
          console.log(`년월: ${idx} (${col})`);
        }
        // 년도 별도
        if (yearIndex === -1 && monthIndex === -1 &&
            (cStr === '년' || cStr === '년도' || cStr === 'year' || cStr === '연도')) {
          yearIndex = idx;
          console.log(`년도: ${idx} (${col})`);
        }
        // 월 별도 (년도 컬럼이 있거나, 단독 '월' 컬럼)
        if (monthIndex === -1 &&
            (cStr === '월' || cStr === '월도' || cStr === 'month' ||
             (cStr.includes('월') && !cStr.includes('사용') && !cStr.includes('요금') && !cStr.includes('금액')))) {
          monthIndex = idx;
          console.log(`월: ${idx} (${col})`);
        }
      });
    }

    // ── 2. 데이터 분석 기반 자동 감지 (헤더 감지 실패 시 보완) ──
    const startRow = headerRowIndex !== -1 ? headerRowIndex + 1 : 0;
    const sampleRows = data.slice(startRow, Math.min(startRow + 6, data.length))
                           .filter(r => r && r.some(c => c != null && c !== ''));

    if ((monthIndex === -1 || usageIndex === -1 || billIndex === -1) && sampleRows.length > 0) {
      console.warn('⚠️ 헤더 감지 보완 – 데이터 값 분석 중');
      const colCount = Math.max(...sampleRows.map(r => r.length));

      for (let ci = 0; ci < colCount; ci++) {
        if (ci === yearIndex) continue;
        const vals = sampleRows.map(r => r ? r[ci] : null).filter(v => v != null && v !== '');
        if (vals.length === 0) continue;

        const hasDate = vals.some(v => {
          const s = String(v);
          return /\d{4}[-년\/\.]\d{1,2}/.test(s) || /^\d{2}[-년\/\.]\d{1,2}/.test(s);
        });
        const numVals = vals.map(v => parseFloat(String(v).replace(/,/g, ''))).filter(n => !isNaN(n));
        const avgNum = numVals.length > 0 ? numVals.reduce((a, b) => a + b, 0) / numVals.length : 0;

        if (hasDate && monthIndex === -1) {
          monthIndex = ci;
          console.log(`📅 날짜 자동 감지: col ${ci}`);
        } else if (numVals.length >= sampleRows.length * 0.5) {
          if (avgNum >= 10000 && billIndex === -1 && ci !== monthIndex) {
            billIndex = ci;
            console.log(`💰 요금 자동 감지: col ${ci} (평균 ${Math.round(avgNum)})`);
          } else if (avgNum > 0 && avgNum < 10000 && usageIndex === -1 && ci !== monthIndex && ci !== billIndex) {
            usageIndex = ci;
            console.log(`⚡ 사용량 자동 감지: col ${ci} (평균 ${Math.round(avgNum)})`);
          }
        }
      }
    }

    // 최종 폴백
    if (headerRowIndex === -1) headerRowIndex = 0;
    if (monthIndex === -1) { monthIndex = 0; console.warn('month → col 0 폴백'); }
    if (usageIndex === -1) { usageIndex = 1; console.warn('usage → col 1 폴백'); }
    if (billIndex === -1)  { billIndex = 2;  console.warn('bill  → col 2 폴백'); }

    console.log('📊 최종 컬럼:', { headerRowIndex, yearIndex, monthIndex, usageIndex, billIndex });

    // ── 3. 데이터 행 파싱 ──
    let lastYear = null;

    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 2) continue;

      const usage = parseFloat(String(row[usageIndex] ?? '').replace(/,/g, ''));
      const bill  = parseFloat(String(row[billIndex]  ?? '').replace(/,/g, ''));

      if (i <= headerRowIndex + 4) {
        console.log(`  행 ${i}:`, { raw: row.slice(0, 12), usage, bill });
      }

      // 청구액만 있어도 허용 (사용량이 없으면 null로 처리)
      if (isNaN(bill)) continue;
      const safeUsage = isNaN(usage) ? null : usage;

      let year = null, month = null;

      // 년도 별도 컬럼
      if (yearIndex !== -1 && row[yearIndex] != null) {
        const y = parseInt(String(row[yearIndex]).replace(/\D/g, ''), 10);
        if (!isNaN(y)) year = y < 100 ? 2000 + y : y;
      }

      const rawDate = row[monthIndex];

      if (typeof rawDate === 'number' && rawDate > 10000) {
        // Excel serial date
        const date = new Date((rawDate - (25567 + 2)) * 86400 * 1000);
        if (!year) year = date.getFullYear();
        month = date.getMonth() + 1;
      } else if (rawDate != null && rawDate !== '') {
        const dStr = String(rawDate).trim();
        // 2024-01 / 2024년1월 / 2024.01
        let match = dStr.match(/(\d{4})[-\.년\/\s]*(\d{1,2})/);
        if (match) {
          if (!year) year = parseInt(match[1], 10);
          month = parseInt(match[2], 10);
        } else {
          // 24-01 / 24년1월
          match = dStr.match(/^(\d{2})[-\.년\/\s]*(\d{1,2})/);
          if (match) {
            if (!year) year = 2000 + parseInt(match[1], 10);
            month = parseInt(match[2], 10);
          } else {
            // 월만 있는 경우 – 이전 행의 년도 이어받기
            match = dStr.match(/^(\d{1,2})\s*월?$/);
            if (match) {
              month = parseInt(match[1], 10);
              if (!year && lastYear) year = lastYear;
            }
          }
        }
      }

      if (year && month && month >= 1 && month <= 12) {
        lastYear = year;
        parsed.push({ year, month, usage: safeUsage, bill });
      }
    }

    console.log(`\n📈 총 ${parsed.length}개월 데이터 파싱됨`);

    if (parsed.length === 0) {
      const preview = data.slice(0, 5)
        .map((r, i) => `${i}: ${r ? r.slice(0, 6).join(' | ') : '(빈 행)'}`)
        .join('\n');
      alert(
        "❌ 데이터 인식 실패\n\n" +
        "유효한 데이터 행을 찾지 못했습니다.\n\n" +
        "📋 파일 상단 내용 (처음 5줄):\n" +
        preview + "\n\n" +
        "💡 해결 방법:\n" +
        "• 날짜 형식: 2024-01, 2024년1월, 24년1월 등\n" +
        "• 컬럼 헤더 포함: 년월/날짜, 사용량/kWh, 요금/청구금액\n" +
        "• 또는 파일 직접 다운로드(.xlsx) 후 업로드"
      );
      return null;
    }

    return parsed;
  };

  // 파싱된 데이터를 state에 반영하는 공통 함수
  const applyParsedData = (parsed, sourceName) => {
    setRawData(parsed);
    const uniqueYears = [...new Set(parsed.map(item => item.year))].sort((a, b) => b - a);
    setYears(uniqueYears);
    if (uniqueYears.length > 0) {
      setStartYear(uniqueYears[0]);
      setStartMonth(1);
    }
    console.log(`✅ [${sourceName}] ${parsed.length}개월 데이터 로드됨`);
  };

  // Parse Excel (로컬 파일)
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
      const parsed = parseSheetData(data, file.name);
      if (!parsed) return;
      applyParsedData(parsed, file.name);
    };
    reader.readAsArrayBuffer(file);
  };

  // 구글 시트 URL → spreadsheetId 추출
  const extractSheetId = (url) => {
    const match = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
  };

  // 구글 시트 불러오기 (CORS 프록시 순차 시도)
  const handleGoogleSheetLoad = async () => {
    const sheetId = extractSheetId(googleSheetUrl);
    if (!sheetId) {
      alert('❌ 올바른 구글 시트 URL이 아닙니다.\n\n구글 시트를 열고 주소창의 URL을 그대로 붙여넣어 주세요.\n예) https://docs.google.com/spreadsheets/d/...');
      return;
    }

    setIsLoadingSheet(true);

    const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv`;
    const gvizUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv`;

    // CORS 우회: 직접 → gviz → 프록시 순서로 시도
    const attempts = [
      { url: csvUrl,                                                          label: '직접' },
      { url: gvizUrl,                                                         label: 'gviz' },
      { url: `https://corsproxy.io/?url=${encodeURIComponent(csvUrl)}`,       label: 'corsproxy' },
      { url: `https://api.allorigins.win/raw?url=${encodeURIComponent(csvUrl)}`, label: 'allorigins' },
    ];

    let csvText = null;

    for (const attempt of attempts) {
      try {
        console.log(`🔄 구글 시트 불러오기 시도 [${attempt.label}]: ${attempt.url.substring(0, 80)}...`);
        const res = await fetch(attempt.url, { signal: AbortSignal.timeout(8000) });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const text = await res.text();
        // HTML 응답(오류 페이지)이 아닌지 확인
        if (text.trim().startsWith('<') && text.includes('<!DOCTYPE')) {
          throw new Error('HTML 오류 페이지 반환');
        }
        csvText = text;
        console.log(`✅ [${attempt.label}] 성공`);
        break;
      } catch (err) {
        console.warn(`⚠️ [${attempt.label}] 실패: ${err.message}`);
      }
    }

    try {
      if (!csvText) throw new Error('모든 경로 실패');

      const wb = XLSX.read(csvText, { type: 'string' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const parsed = parseSheetData(data, '구글 시트');
      if (!parsed) return;

      setFileName('구글 시트에서 불러옴');
      applyParsedData(parsed, '구글 시트');
    } catch (err) {
      console.error(err);
      alert(
        '❌ 구글 시트를 불러오지 못했습니다.\n\n' +
        '✅ 해결 방법:\n' +
        '1. 구글 시트 상단 [공유] 버튼 클릭\n' +
        '2. "링크가 있는 모든 사용자" 선택 후 권한을 "뷰어"로 설정\n' +
        '3. 링크 복사 후 다시 붙여넣기\n\n' +
        '💡 또는 파일을 직접 다운로드(.xlsx)하여 업로드해 주세요.'
      );
    } finally {
      setIsLoadingSheet(false);
    }
  };

  // Generate Monthly Template
  const monthlyData = useMemo(() => {
    if (!startYear || !startMonth || rawData.length === 0) return [];
    
    let res = [];
    let currentY = Number(startYear);
    let currentM = Number(startMonth);

    for (let c = 0; c < 12; c++) {
      const found = rawData.find(d => d.year === currentY && d.month === currentM);
      const rawUsage = found ? found.usage : null;   // null = 데이터 없거나 시트에 사용량 컬럼 없음
      const usage = rawUsage ?? 0;                   // 계산용은 0으로 대체
      const totalBill = found ? found.bill : 0;

      const safeOtherPowerBill = Number(otherPowerBill) || 0;
      const safeSavingsRate = Number(savingsRate) || 0;

      let boilerBill = 0;

      if (calculationMethod === 'billing') {
        // 방식 1: 청구서 기반 (기저부하 차감)
        boilerBill = Math.max(0, totalBill - safeOtherPowerBill);
      } else if (calculationMethod === 'theoretical') {
        // 방식 2: 이론 계산 (용량 × 시간 × 일수)
        const capacity = Number(boilerCapacity) || 0;
        const hoursPerDay = Number(dailyUsageHours) || 0;
        const daysInYear = Number(annualOperatingDays) || 365;
        const daysPerMonth = daysInYear / 12; // 평균 월 가동일

        // 보일러 사용량 (kWh) = 용량(kW) × 시간/일 × 일수
        let boilerUsage = capacity * hoursPerDay * daysPerMonth;

        // ⚠️ 안전장치: 이론 계산 보일러 사용량이 실제 전체 사용량을 초과하지 않도록 제한
        if (usage > 0 && boilerUsage > usage) {
          boilerUsage = usage * 0.95; // 최대 전체 사용량의 95%까지만 허용
        }

        // 평균 전기단가 계산 (원/kWh)
        const averageRate = usage > 0 ? totalBill / usage : getElectricityRate(); // 계약종별 기본값

        // 보일러 요금 = 보일러 사용량 × 전기단가
        boilerBill = boilerUsage * averageRate;
      }

      const savings = boilerBill * (safeSavingsRate / 100);
      const newBoilerBill = boilerBill - savings;

      res.push({
        displayMonth: `${currentY.toString().slice(-2)}년 ${currentM}월`,
        year: currentY,
        month: currentM,
        usage,
        rawUsage,   // null이면 시트에 사용량 데이터 없음
        totalBill,
        boilerBill,
        expectedSavings: savings,
        newBoilerBill,
        newTotalBill: newBoilerBill + (totalBill > 0 ? safeOtherPowerBill : 0)
      });

      currentM++;
      if (currentM > 12) {
        currentM = 1;
        currentY++;
      }
    }
    return res;
  }, [rawData, startYear, startMonth, otherPowerBill, savingsRate, calculationMethod, boilerCapacity, dailyUsageHours, annualOperatingDays, contractType, customRate, rateSource]);

  const facilityInvestment = 
    (Number(boilerCount)||0) * (Number(boilerUnitPrice)||0) + 
    (Number(boilerCount)||0) * (Number(installationCost)||0) + 
    (Number(thermalTankCost)||0) + 
    (Number(electricalWorkCost)||0) + 
    (Number(otherCosts)||0);

  const kepcoCost = facilityInvestment * ((Number(kepcoRate)||0) / 100);
  const sgiCost = facilityInvestment * ((Number(sgiRate)||0) / 100);
  const totalCost = facilityInvestment + kepcoCost + sgiCost;

  // Aggregated Stats
  const stats = useMemo(() => {
    if (!monthlyData.length) return null;
    let totalUsage = 0, totalBill = 0, totalBoilerBill = 0, totalSavings = 0, totalNewBill = 0;
    let usageAvailable = false;

    monthlyData.forEach(d => {
      totalUsage += d.usage;
      if (d.rawUsage !== null) usageAvailable = true;
      totalBill += d.totalBill;
      totalBoilerBill += d.boilerBill;
      totalSavings += d.expectedSavings;
      totalNewBill += d.newTotalBill;
    });

    const n = monthlyData.length;
    const monthlyInstallment = (Number(installmentMonths)||0) > 0 ? totalCost / Number(installmentMonths) : 0;
    const monthlyRental = monthlyInstallment;

    const avgTotalBill = totalBill / n;             // 개선 전 월 평균 전기료
    const avgNewTotalBill = totalNewBill / n;        // 개선 후 월 평균 전기료
    const monthlyTotalBurden = avgNewTotalBill + monthlyInstallment; // 개선 후 실제 월 부담 (전기료+임대료)

    const netBenefitMonthly = (totalSavings / 12) - monthlyInstallment;

    // 손익분기점(BEP) 절감율 계산 (%)
    const breakevenRate = (totalBoilerBill > 0 && Number(installmentMonths) > 0)
      ? (totalCost * 1200) / (totalBoilerBill * Number(installmentMonths))
      : 0;

    return {
      totalUsage,
      totalBill,
      totalBoilerBill,
      totalSavings,
      monthlyInstallment,
      monthlyRental,
      netBenefitMonthly,
      averageSavingsMonthly: totalSavings / 12,
      avgTotalBill,
      avgNewTotalBill,
      monthlyTotalBurden,
      usageAvailable,
      breakevenRate
    };
  }, [monthlyData, totalCost, installmentMonths]);

  // Download Image (Infographic)
  const handleDownloadImage = async () => {
    if (!infographicRef.current) return;
    try {
      const el = infographicRef.current;
      const canvas = await html2canvas(el, {
        scale: 2,
        backgroundColor: '#0f172a',
        useCORS: true,
      });
      const dataUrl = canvas.toDataURL('image/png');
      const link = document.createElement('a');
      link.download = `고효율보일러_대시보드_${startYear}년${startMonth}월.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      alert("이미지 저장 중 오류가 발생했습니다.");
      console.error(err);
    }
  };

  // Download Excel
  const handleDownloadExcel = () => {
    if (monthlyData.length === 0) return;
    
    const wsData = monthlyData.map(d => ({
      '월': d.displayMonth,
      '전기 사용량 (kWh)': d.usage,
      '기존 총 요금 (원)': d.totalBill,
      '기존 보일러 요금 (원)': d.boilerBill,
      '예상 절감액 (원)': d.expectedSavings,
      '새 보일러 요금 (원)': d.newBoilerBill,
      '개선 후 총 요금 (원)': d.newTotalBill
    }));

    // Summary Rows
    wsData.push({});
    wsData.push({
      '월': '연간 총계',
      '기존 총 요금 (원)': stats.totalBill,
      '기존 보일러 요금 (원)': stats.totalBoilerBill,
      '예상 절감액 (원)': stats.totalSavings,
      '새 보일러 요금 (원)': stats.totalBoilerBill - stats.totalSavings
    });
    wsData.push({
      '월': '월 할부/임대료',
      '기존 총 요금 (원)': stats.monthlyInstallment
    });
    wsData.push({
      '월': '월 평균 순이익 (절감액 - 할부금)',
      '기존 총 요금 (원)': stats.netBenefitMonthly
    });

    const ws = XLSX.utils.json_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `장기분석보고서`);
    XLSX.writeFile(wb, `고효율보일러_제안분석_${startYear}년${startMonth}월기준.xlsx`);
  };

  return (
    <div className="container fade-in">
      <header className="header">
        <h1>고효율 보일러 에너지 진단 & 제안</h1>
        
      </header>

      {/* 1. 데이터 업로드 및 기본 설정 */}
      <section className="glass-panel mb-6">
        <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
          <FileSpreadsheet size={24} color="var(--primary-color)" />
          데이터 입력 방법 선택
        </h2>

        {/* 탭 선택 */}
        <div className="tab-container">
          <button
            onClick={() => setInputMode('excel')}
            className={`tab-button ${inputMode === 'excel' ? 'active' : ''}`}
          >
            📊 엑셀 파일 업로드
          </button>
          <button
            onClick={() => setInputMode('manual')}
            className={`tab-button ${inputMode === 'manual' ? 'active' : ''}`}
          >
            ✍️ 수동 입력
          </button>
        </div>

        {/* 엑셀 업로드 모드 */}
        {inputMode === 'excel' && (
          <div className="grid-2">
            <div style={{display: 'flex', flexDirection: 'column', gap: '16px'}}>
              <div className="file-upload-wrapper">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={handleFileUpload}
                  className="file-upload-input"
                  ref={fileInputRef}
                />
                <div className="file-upload-label">
                  <Upload size={32} color="var(--text-secondary)" style={{marginBottom: '12px'}} />
                  <div style={{fontWeight: 600}}>엑셀/구글시트 원본 데이터 업로드</div>
                  <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)', marginTop: '4px'}}>
                    {fileName ? `업로드됨: ${fileName}` : '클릭하거나 파일을 드래그하세요'}
                  </div>
                </div>
              </div>

              {/* 구글 시트 URL 입력 */}
              <div style={{display: 'flex', alignItems: 'center', gap: '12px', color: 'var(--text-secondary)', fontSize: '0.85rem'}}>
                <div style={{flex: 1, height: '1px', background: 'rgba(255,255,255,0.1)'}} />
                <span>또는 구글 시트 URL 입력</span>
                <div style={{flex: 1, height: '1px', background: 'rgba(255,255,255,0.1)'}} />
              </div>

              <div style={{display: 'flex', gap: '8px'}}>
                <input
                  type="text"
                  className="form-control"
                  value={googleSheetUrl}
                  onChange={e => setGoogleSheetUrl(e.target.value)}
                  placeholder="https://docs.google.com/spreadsheets/d/..."
                  style={{flex: 1, fontSize: '0.85rem'}}
                  onKeyDown={e => e.key === 'Enter' && handleGoogleSheetLoad()}
                />
                <button
                  className="btn btn-accent"
                  onClick={handleGoogleSheetLoad}
                  disabled={isLoadingSheet || !googleSheetUrl.trim()}
                  style={{whiteSpace: 'nowrap', opacity: (!googleSheetUrl.trim() || isLoadingSheet) ? 0.5 : 1}}
                >
                  {isLoadingSheet ? '불러오는 중...' : '불러오기'}
                </button>
              </div>

              <div style={{padding: '12px', borderRadius: '8px', background: 'rgba(16, 185, 129, 0.08)', border: '1px solid rgba(16, 185, 129, 0.2)', fontSize: '0.82rem', color: 'var(--text-secondary)'}}>
                <strong style={{color: 'var(--success-color)'}}>📋 구글 시트 사용 조건:</strong>
                {' '}구글 시트 공유 설정이 <strong>"링크가 있는 모든 사용자 - 뷰어"</strong> 이상으로 설정되어 있어야 합니다.
              </div>

              <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', fontSize: '0.85rem', color: 'var(--text-secondary)'}}>
                <div style={{fontWeight: '600', color: 'var(--primary-color)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px'}}>
                  <span style={{fontSize: '1rem'}}>💡</span> 데이터 업로드 필수사항
                </div>
                <ul style={{margin: 0, paddingLeft: '20px', display: 'flex', flexDirection: 'column', gap: '6px', lineHeight: '1.4'}}>
                  <li>각 데이터 열(차트) 이름에 <strong>[월/년/일]</strong>, <strong>[사용량/전력/kWh]</strong>, <strong>[요금/금액/청구]</strong> 단어가 포함되어 있어야 정상 매핑됩니다.</li>
                  <li>문서 상단에 데이터와 관계없는 불필요한 내용이나 텍스트가 너무 길면(20줄 이상) 표를 자동 인식하지 못할 수 있습니다.</li>
                  <li>계속해서 업로드 에러가 난다면, 단순히 열 이름을 <strong>"년월", "사용량", "요금"</strong>으로 알맞게 수정한 후 다시 업로드해 보세요.</li>
                </ul>
              </div>
            </div>

            <div>
              <div className="grid-2">
                <div className="form-group">
                  <label>조회 시작 연도</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="number" className="form-control" value={startYear} onChange={e => setStartYear(e.target.value)} placeholder="예: 2025" />
                    <span>년</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>조회 시작 월 (입력 월부터 1년)</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="number" className="form-control" value={startMonth} onChange={e => setStartMonth(e.target.value)} placeholder="예: 3" min="1" max="12" />
                    <span>월</span>
                  </div>
                </div>
              </div>
              {monthlyData.length > 0 && (
                <>
                  <h3 style={{marginTop: '32px', marginBottom: '16px', fontSize: '1rem', color: 'var(--text-primary)', borderTop: '1px solid rgba(255,255,255,0.05)', paddingTop: '24px'}}>
                    보일러 사용량 계산 방식 선택
                  </h3>

                  {/* 계산 방식 선택 */}
                  <div style={{display: 'flex', gap: '12px', marginBottom: '24px'}}>
                    <button
                      onClick={() => setCalculationMethod('billing')}
                      className={`input-mode-button ${calculationMethod === 'billing' ? 'active' : ''}`}
                    >
                      <span style={{fontSize: '1.5rem'}}>📋</span>
                      <span>청구서 기반 계산</span>
                      <span style={{fontSize: '0.8rem', opacity: 0.8}}>전체 요금 - 기저부하</span>
                    </button>
                    <button
                      onClick={() => setCalculationMethod('theoretical')}
                      className={`input-mode-button ${calculationMethod === 'theoretical' ? 'active' : ''}`}
                    >
                      <span style={{fontSize: '1.5rem'}}>🔬</span>
                      <span>이론 계산</span>
                      <span style={{fontSize: '0.8rem', opacity: 0.8}}>용량 × 시간 × 일수</span>
                    </button>
                  </div>

                  {/* 청구서 기반 계산 입력 */}
                  {calculationMethod === 'billing' && (
                    <div className="form-group">
                      <label>월별 비보일러 전기요금 (기저부하)</label>
                      <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                        <input
                          type="text"
                          className="form-control"
                          value={formatInput(otherPowerBill)}
                          onChange={handleMoneyChange(setOtherPowerBill)}
                          placeholder="0"
                        />
                        <span>원/월</span>
                      </div>
                      <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '8px'}}>
                        💡 전체 전기요금에서 이 금액을 뺀 나머지가 "순수 보일러 전기요금"으로 산정됩니다.
                      </div>
                    </div>
                  )}

                  {/* 이론 계산 입력 */}
                  {calculationMethod === 'theoretical' && (
                    <>
                      <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '16px'}}>
                        <div style={{fontWeight: '600', color: 'var(--primary-color)', marginBottom: '8px'}}>
                          🔬 이론적 계산 방식
                        </div>
                        <p style={{margin: 0}}>보일러 사용량(kWh) = 전기용량(kW) × 일평균 사용시간(시간/일) × 월평균 가동일(일/월)</p>
                        <p style={{margin: '4px 0 0 0'}}>보일러 요금 = 보일러 사용량 × 평균 전기단가</p>
                      </div>

                      {/* 이론 계산값이 비정상적으로 큰 경우 경고 */}
                      {(() => {
                        const capacity = Number(boilerCapacity) || 0;
                        const hoursPerDay = Number(dailyUsageHours) || 0;
                        const daysPerMonth = (Number(annualOperatingDays) || 365) / 12;
                        const theoreticalUsage = capacity * hoursPerDay * daysPerMonth;
                        const avgUsage = rawData.length > 0 ? rawData.reduce((sum, d) => sum + d.usage, 0) / rawData.length : 0;

                        if (theoreticalUsage > avgUsage && avgUsage > 0) {
                          return (
                            <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '16px'}}>
                              <div style={{fontWeight: '600', color: 'rgb(239, 68, 68)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px'}}>
                                <span style={{fontSize: '1rem'}}>⚠️</span> 입력값 검토 필요
                              </div>
                              <p style={{margin: 0}}>
                                이론 계산된 보일러 사용량 ({Math.round(theoreticalUsage).toLocaleString()} kWh)이 실제 월평균 전체 사용량 ({Math.round(avgUsage).toLocaleString()} kWh)보다 {Math.round(theoreticalUsage / avgUsage * 10) / 10}배 높습니다.
                              </p>
                              <p style={{margin: '8px 0 0 0'}}>
                                💡 <strong>전기용량, 일평균 사용시간, 가동일수를 다시 확인</strong>해주세요. 입력값이 너무 크면 부정확한 결과가 나올 수 있습니다.
                              </p>
                            </div>
                          );
                        }
                        return null;
                      })()}

                      <div className="grid-3">
                        <div className="form-group">
                          <label>보일러 전기용량</label>
                          <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                            <input
                              type="number"
                              className="form-control"
                              value={boilerCapacity}
                              onChange={e => setBoilerCapacity(e.target.value)}
                              placeholder="0"
                              step="0.1"
                            />
                            <span>kW</span>
                          </div>
                        </div>
                        <div className="form-group">
                          <label>일평균 사용시간</label>
                          <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                            <input
                              type="number"
                              className="form-control"
                              value={dailyUsageHours}
                              onChange={e => setDailyUsageHours(e.target.value)}
                              placeholder="0"
                              step="0.1"
                            />
                            <span>시간/일</span>
                          </div>
                        </div>
                        <div className="form-group">
                          <label>년간 가동일</label>
                          <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                            <input
                              type="number"
                              className="form-control"
                              value={annualOperatingDays}
                              onChange={e => setAnnualOperatingDays(e.target.value)}
                              placeholder="365"
                            />
                            <span>일/년</span>
                          </div>
                        </div>
                      </div>
                    </>
                  )}
                </>
              )}
            </div>
          </div>
        )}

        {/* 수동 입력 모드 */}
        {inputMode === 'manual' && (
          <div>
            <div className="grid-2" style={{marginBottom: '24px'}}>
              <div className="form-group">
                <label>기준 연도</label>
                <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                  <input
                    type="number"
                    className="form-control"
                    value={startYear}
                    onChange={e => setStartYear(e.target.value)}
                    placeholder="예: 2025"
                  />
                  <span>년</span>
                </div>
              </div>
              <div className="form-group">
                <label>조회 시작 월</label>
                <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                  <input
                    type="number"
                    className="form-control"
                    value={startMonth}
                    onChange={e => setStartMonth(e.target.value)}
                    placeholder="예: 1"
                    min="1"
                    max="12"
                  />
                  <span>월</span>
                </div>
              </div>
            </div>

            {/* 입력 방식 선택 */}
            <div style={{display: 'flex', gap: '12px', marginBottom: '24px'}}>
              <button
                onClick={() => setManualInputType('annual')}
                className={`input-mode-button ${manualInputType === 'annual' ? 'active' : ''}`}
              >
                <span style={{fontSize: '1.5rem'}}>⚡</span>
                <span>연간 총액 입력 (빠른 입력)</span>
                <span style={{fontSize: '0.8rem', opacity: 0.8}}>고객에게 빠르게 보여주기</span>
              </button>
              <button
                onClick={() => setManualInputType('monthly')}
                className={`input-mode-button ${manualInputType === 'monthly' ? 'active' : ''}`}
              >
                <span style={{fontSize: '1.5rem'}}>📅</span>
                <span>월별 개별 입력 (상세 입력)</span>
                <span style={{fontSize: '0.8rem', opacity: 0.8}}>정확한 월별 데이터</span>
              </button>
            </div>

            {/* 연간 총액 입력 */}
            {manualInputType === 'annual' && (
              <>
                <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '24px'}}>
                  <div style={{fontWeight: '600', color: 'var(--primary-color)', marginBottom: '8px'}}>
                    💡 연간 전기 요금만 입력하세요
                  </div>
                  <p style={{margin: 0}}>입력한 요금이 12개월로 균등 분배되며, 사용량은 자동으로 추정됩니다. 빠른 시뮬레이션에 적합합니다.</p>
                </div>

                <div style={{padding: '32px', background: 'rgba(30, 41, 59, 0.5)', borderRadius: '12px', border: '1px solid var(--card-border)', marginBottom: '16px'}}>
                  <div className="form-group" style={{maxWidth: '600px', margin: '0 auto'}}>
                    <label style={{fontSize: '1.2rem', fontWeight: 700, marginBottom: '16px', display: 'block', textAlign: 'center'}}>연간 총 전기 요금</label>
                    <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                      <input
                        type="text"
                        className="form-control"
                        value={formatInput(annualBill)}
                        onChange={handleMoneyChange(setAnnualBill)}
                        placeholder="예: 20000000"
                        style={{fontSize: '1.5rem', padding: '20px', textAlign: 'center', fontWeight: 'bold'}}
                      />
                      <span style={{fontWeight: 'bold', fontSize: '1.2rem', whiteSpace: 'nowrap'}}>원 / 년</span>
                    </div>
                    <div style={{fontSize: '0.9rem', color: 'var(--text-secondary)', marginTop: '16px', textAlign: 'center', padding: '12px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '8px'}}>
                      <div style={{marginBottom: '8px'}}>
                        📊 <strong>월평균 요금:</strong> {annualBill ? formatMoney(Math.round(Number(annualBill) / 12)) : '0'} 원
                      </div>
                      <div style={{fontSize: '0.8rem', opacity: 0.8}}>
                        ⚡ 사용량은 선택한 계약종별의 평균 단가({getElectricityRate()}원/kWh)를 기준으로 자동 추정됩니다
                      </div>
                    </div>
                  </div>
                </div>

                <button
                  className="btn btn-accent"
                  onClick={applyAnnualData}
                  style={{width: '100%', padding: '16px', fontSize: '1.1rem'}}
                >
                  <Calculator size={20} /> 연간 데이터 적용 및 분석 시작
                </button>
              </>
            )}

            {/* 월별 개별 입력 */}
            {manualInputType === 'monthly' && (
              <>
                <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '16px'}}>
                  <div style={{fontWeight: '600', color: 'var(--primary-color)', marginBottom: '8px'}}>
                    💡 12개월치 전기 사용량과 요금을 입력하세요
                  </div>
                  <p style={{margin: 0}}>입력하지 않은 월은 자동으로 0으로 처리됩니다.</p>
                </div>

                <div style={{maxHeight: '400px', overflowY: 'auto', marginBottom: '16px'}}>
                  <div className="monthly-input-grid">
                    {manualData.map((data, index) => (
                      <div key={index} style={{padding: '16px', background: 'rgba(30, 41, 59, 0.5)', borderRadius: '8px', border: '1px solid var(--card-border)'}}>
                        <div style={{fontWeight: 'bold', marginBottom: '12px', color: 'var(--primary-color)'}}>{data.month}월</div>
                        <div className="form-group" style={{marginBottom: '12px'}}>
                          <label style={{fontSize: '0.85rem'}}>사용량 (kWh)</label>
                          <input
                            type="text"
                            className="form-control"
                            value={formatInput(data.usage)}
                            onChange={e => handleManualDataChange(index, 'usage', e.target.value)}
                            placeholder="0"
                          />
                        </div>
                        <div className="form-group">
                          <label style={{fontSize: '0.85rem'}}>청구 요금 (원)</label>
                          <input
                            type="text"
                            className="form-control"
                            value={formatInput(data.bill)}
                            onChange={e => handleManualDataChange(index, 'bill', e.target.value)}
                            placeholder="0"
                          />
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                <button
                  className="btn btn-accent"
                  onClick={applyManualData}
                  style={{width: '100%', padding: '16px', fontSize: '1.1rem'}}
                >
                  <Calculator size={20} /> 월별 데이터 적용 및 분석 시작
                </button>
              </>
            )}
          </div>
        )}
      </section>

      {/* 전기요금 계약종별 선택 */}
      {(rawData.length > 0 || (inputMode === 'manual' && manualInputType === 'annual' && annualBill)) && (
        <section className="glass-panel mb-6">
          <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
            <Zap size={24} color="var(--warning-color)" />
            전기요금 단가 설정
            {getActualAverageRate() === null && (
              <span style={{fontSize: '0.85rem', fontWeight: 'normal', color: 'var(--warning-color)', marginLeft: '8px'}}>
                (필수)
              </span>
            )}
          </h2>

          <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '24px'}}>
            <div style={{fontWeight: '600', color: 'var(--primary-color)', marginBottom: '8px'}}>
              💡 한국전력공사 2026년 4월 16일 시행 요금표 기준
            </div>
            {getActualAverageRate() !== null ? (
              <>
                <p style={{margin: '0 0 16px 0'}}>
                  Excel 데이터에서 실제 평균 단가를 계산했습니다. 아래에서 사용할 전기단가를 선택하세요.
                </p>

                {/* 전기단가 선택 라디오 버튼 */}
                <div style={{display: 'flex', flexDirection: 'column', gap: '12px'}}>
                  <label
                    style={{
                      padding: '16px',
                      borderRadius: '8px',
                      border: rateSource === 'actual' ? '2px solid var(--primary-color)' : '2px solid rgba(255,255,255,0.1)',
                      background: rateSource === 'actual' ? 'rgba(59, 130, 246, 0.15)' : 'rgba(30, 41, 59, 0.3)',
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '12px',
                      transition: 'all 0.2s'
                    }}
                  >
                    <input
                      type="radio"
                      name="rateSource"
                      value="actual"
                      checked={rateSource === 'actual'}
                      onChange={(e) => setRateSource(e.target.value)}
                      style={{width: '18px', height: '18px', cursor: 'pointer'}}
                    />
                    <div style={{flex: 1}}>
                      <div style={{fontWeight: 600, color: 'var(--text-primary)', marginBottom: '4px'}}>
                        📊 실제 데이터 사용 (추천)
                      </div>
                      <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)'}}>
                        업로드된 Excel 파일의 평균 단가: <strong style={{color: 'var(--primary-color)'}}>{getActualAverageRate()}원/kWh</strong>
                      </div>
                    </div>
                  </label>

                  <label
                    style={{
                      padding: '16px',
                      borderRadius: '8px',
                      border: rateSource === 'contract' ? '2px solid var(--primary-color)' : '2px solid rgba(255,255,255,0.1)',
                      background: rateSource === 'contract' ? 'rgba(59, 130, 246, 0.15)' : 'rgba(30, 41, 59, 0.3)',
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '12px',
                      transition: 'all 0.2s'
                    }}
                  >
                    <input
                      type="radio"
                      name="rateSource"
                      value="contract"
                      checked={rateSource === 'contract'}
                      onChange={(e) => setRateSource(e.target.value)}
                      style={{width: '18px', height: '18px', cursor: 'pointer'}}
                    />
                    <div style={{flex: 1}}>
                      <div style={{fontWeight: 600, color: 'var(--text-primary)', marginBottom: '4px'}}>
                        📋 계약종별 단가 사용
                      </div>
                      <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)'}}>
                        선택한 계약종별의 평균 단가: <strong style={{color: 'var(--warning-color)'}}>
                          {contractType === 'custom'
                            ? `${Number(customRate) || 120}원/kWh (직접입력)`
                            : `${contractTypes[contractType]?.rate}원/kWh (${contractTypes[contractType]?.name})`
                          }
                        </strong>
                      </div>
                    </div>
                  </label>
                </div>
              </>
            ) : (
              <p style={{margin: 0}}>
                계약종별을 선택하면 해당 종별의 평균 전기단가가 자동 적용됩니다. 정확한 시뮬레이션을 위해 실제 계약종별을 선택해주세요.
              </p>
            )}
          </div>

          <div className="grid-2" style={{gap: '16px'}}>
            {Object.entries(contractTypes).filter(([key]) => key !== 'custom').map(([key, info]) => (
              <label
                key={key}
                style={{
                  padding: '20px',
                  borderRadius: '12px',
                  border: contractType === key ? '2px solid var(--primary-color)' : '2px solid var(--card-border)',
                  background: contractType === key ? 'rgba(59, 130, 246, 0.1)' : 'rgba(30, 41, 59, 0.5)',
                  cursor: 'pointer',
                  transition: 'all 0.2s',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '12px'
                }}
                onMouseEnter={(e) => {
                  if (contractType !== key) {
                    e.currentTarget.style.borderColor = 'rgba(59, 130, 246, 0.5)';
                    e.currentTarget.style.background = 'rgba(30, 41, 59, 0.7)';
                  }
                }}
                onMouseLeave={(e) => {
                  if (contractType !== key) {
                    e.currentTarget.style.borderColor = 'var(--card-border)';
                    e.currentTarget.style.background = 'rgba(30, 41, 59, 0.5)';
                  }
                }}
              >
                <input
                  type="radio"
                  name="contractType"
                  value={key}
                  checked={contractType === key}
                  onChange={(e) => setContractType(e.target.value)}
                  style={{width: '20px', height: '20px', cursor: 'pointer'}}
                />
                <div style={{flex: 1}}>
                  <div style={{fontWeight: 600, fontSize: '1.05rem', marginBottom: '4px', color: 'var(--text-primary)'}}>
                    {info.name}
                  </div>
                  <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '6px'}}>
                    {info.description}
                  </div>
                  <div style={{fontSize: '1.1rem', fontWeight: 700, color: 'var(--primary-color)'}}>
                    평균 {info.rate}원/kWh
                  </div>
                </div>
              </label>
            ))}

            {/* 직접입력 옵션 */}
            <label
              style={{
                padding: '20px',
                borderRadius: '12px',
                border: contractType === 'custom' ? '2px solid var(--primary-color)' : '2px solid var(--card-border)',
                background: contractType === 'custom' ? 'rgba(59, 130, 246, 0.1)' : 'rgba(30, 41, 59, 0.5)',
                cursor: 'pointer',
                transition: 'all 0.2s',
                display: 'flex',
                flexDirection: 'column',
                gap: '12px'
              }}
              onMouseEnter={(e) => {
                if (contractType !== 'custom') {
                  e.currentTarget.style.borderColor = 'rgba(59, 130, 246, 0.5)';
                  e.currentTarget.style.background = 'rgba(30, 41, 59, 0.7)';
                }
              }}
              onMouseLeave={(e) => {
                if (contractType !== 'custom') {
                  e.currentTarget.style.borderColor = 'var(--card-border)';
                  e.currentTarget.style.background = 'rgba(30, 41, 59, 0.5)';
                }
              }}
            >
              <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                <input
                  type="radio"
                  name="contractType"
                  value="custom"
                  checked={contractType === 'custom'}
                  onChange={(e) => setContractType(e.target.value)}
                  style={{width: '20px', height: '20px', cursor: 'pointer'}}
                />
                <div style={{flex: 1}}>
                  <div style={{fontWeight: 600, fontSize: '1.05rem', marginBottom: '4px', color: 'var(--text-primary)'}}>
                    직접입력
                  </div>
                  <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)'}}>
                    사용자 지정 전기단가
                  </div>
                </div>
              </div>
              {contractType === 'custom' && (
                <div style={{display: 'flex', alignItems: 'center', gap: '8px', marginTop: '8px'}}>
                  <input
                    type="number"
                    className="form-control"
                    value={customRate}
                    onChange={(e) => setCustomRate(e.target.value)}
                    placeholder="예: 120"
                    step="0.1"
                    style={{flex: 1}}
                  />
                  <span style={{whiteSpace: 'nowrap', fontWeight: 600}}>원/kWh</span>
                </div>
              )}
            </label>
          </div>

          <div style={{marginTop: '24px', padding: '16px', borderRadius: '8px', background: 'rgba(255, 193, 7, 0.1)', border: '1px solid rgba(255, 193, 7, 0.3)'}}>
            <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
              <div style={{fontSize: '1.5rem'}}>⚡</div>
              <div style={{flex: 1}}>
                <div style={{fontWeight: 600, color: 'var(--text-primary)', marginBottom: '4px'}}>
                  현재 적용 중인 전기단가
                </div>
                <div style={{fontSize: '1.5rem', fontWeight: 700, color: 'var(--warning-color)'}}>
                  {getElectricityRate()}원/kWh
                </div>
                <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px'}}>
                  {getActualAverageRate() !== null ? (
                    <>
                      {rateSource === 'actual' ? (
                        <>
                          📊 <strong style={{color: 'var(--primary-color)'}}>실제 데이터 사용 중</strong> (업로드된 Excel 파일의 평균 단가)
                        </>
                      ) : (
                        <>
                          📋 <strong style={{color: 'var(--warning-color)'}}>계약종별 단가 사용 중</strong> ({contractTypes[contractType]?.name || '직접입력'})
                          {(() => {
                            const expectedRate = contractType === 'custom'
                              ? Number(customRate) || 120
                              : contractTypes[contractType]?.rate || 120;
                            const actualRate = getActualAverageRate();
                            const deviation = Math.abs(actualRate - expectedRate) / expectedRate;

                            if (deviation > 0.3) {
                              return (
                                <div style={{marginTop: '8px', padding: '8px', background: 'rgba(239, 68, 68, 0.1)', borderRadius: '4px', border: '1px solid rgba(239, 68, 68, 0.2)'}}>
                                  ⚠️ 실제 데이터의 평균 단가({actualRate}원/kWh)와 {Math.round(deviation * 100)}% 차이가 있습니다.
                                </div>
                              );
                            }
                            return null;
                          })()}
                        </>
                      )}
                    </>
                  ) : (
                    <>
                      {contractTypes[contractType]?.name || '직접입력'} ({contractTypes[contractType]?.description || '사용자 지정 단가'})
                    </>
                  )}
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {/* 비보일러 전기요금 입력 (수동 입력 모드일 때만 별도 섹션으로) */}
      {inputMode === 'manual' && monthlyData.length > 0 && (
        <section className="glass-panel mb-6">
          <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
            <Settings size={20} color="var(--primary-color)" />
            보일러 사용량 계산 방식
          </h2>

          {/* 계산 방식 선택 */}
          <div style={{display: 'flex', gap: '12px', marginBottom: '24px'}}>
            <button
              onClick={() => setCalculationMethod('billing')}
              className={`input-mode-button ${calculationMethod === 'billing' ? 'active' : ''}`}
            >
              <span style={{fontSize: '1.5rem'}}>📋</span>
              <span>청구서 기반 계산</span>
              <span style={{fontSize: '0.8rem', opacity: 0.8}}>전체 요금 - 기저부하</span>
            </button>
            <button
              onClick={() => setCalculationMethod('theoretical')}
              className={`input-mode-button ${calculationMethod === 'theoretical' ? 'active' : ''}`}
            >
              <span style={{fontSize: '1.5rem'}}>🔬</span>
              <span>이론 계산</span>
              <span style={{fontSize: '0.8rem', opacity: 0.8}}>용량 × 시간 × 일수</span>
            </button>
          </div>

          {/* 청구서 기반 계산 입력 */}
          {calculationMethod === 'billing' && (
            <div className="form-group">
              <label>월별 비보일러 전기요금 (기저부하)</label>
              <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                <input
                  type="text"
                  className="form-control"
                  value={formatInput(otherPowerBill)}
                  onChange={handleMoneyChange(setOtherPowerBill)}
                  placeholder="0"
                />
                <span>원/월</span>
              </div>
              <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '8px'}}>
                💡 전체 전기요금에서 이 금액을 뺀 나머지가 "순수 보일러 전기요금"으로 산정됩니다.
              </div>

              {/* Excel 데이터 이상 감지 */}
              {(() => {
                const avgBill = rawData.length > 0 ? rawData.reduce((sum, d) => sum + d.bill, 0) / rawData.length : 0;
                const avgUsage = rawData.length > 0 ? rawData.reduce((sum, d) => sum + d.usage, 0) / rawData.length : 0;
                const avgRate = avgUsage > 0 ? avgBill / avgUsage : 0;

                // 평균 전기단가가 명백히 비정상적인 경우에만 경고 (10원 미만 또는 500원 초과)
                if ((avgRate < 10 && avgRate > 0) || avgRate > 500) {
                  return (
                    <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)', fontSize: '0.85rem', marginTop: '16px'}}>
                      <div style={{fontWeight: '600', color: 'rgb(239, 68, 68)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px'}}>
                        <span style={{fontSize: '1rem'}}>⚠️</span> Excel 데이터 파싱 오류 가능성
                      </div>
                      <p style={{margin: 0, color: 'var(--text-secondary)'}}>
                        업로드된 데이터의 평균 전기단가가 <strong>{avgRate.toFixed(2)}원/kWh</strong>로 비정상적입니다. (정상 범위: 10~500원/kWh)
                      </p>
                      <p style={{margin: '8px 0 0 0', color: 'var(--text-secondary)'}}>
                        💡 Excel 파일의 <strong>"요금" 컬럼이 잘못 인식</strong>되었을 가능성이 높습니다. 브라우저 개발자 도구(F12) → Console에서 "요금 컬럼" 로그를 확인하시거나, Excel 파일의 컬럼명을 "년월", "사용량", "청구금액"으로 명확히 수정해주세요.
                      </p>
                    </div>
                  );
                }
                return null;
              })()}
            </div>
          )}

          {/* 이론 계산 입력 */}
          {calculationMethod === 'theoretical' && (
            <>
              <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '16px'}}>
                <div style={{fontWeight: '600', color: 'var(--primary-color)', marginBottom: '8px'}}>
                  🔬 이론적 계산 방식
                </div>
                <p style={{margin: 0}}>보일러 사용량(kWh) = 전기용량(kW) × 일평균 사용시간(시간/일) × 월평균 가동일(일/월)</p>
                <p style={{margin: '4px 0 0 0'}}>보일러 요금 = 보일러 사용량 × 평균 전기단가</p>
              </div>

              {/* 이론 계산값이 비정상적으로 큰 경우 경고 */}
              {(() => {
                const capacity = Number(boilerCapacity) || 0;
                const hoursPerDay = Number(dailyUsageHours) || 0;
                const daysPerMonth = (Number(annualOperatingDays) || 365) / 12;
                const theoreticalUsage = capacity * hoursPerDay * daysPerMonth;
                const avgUsage = rawData.length > 0 ? rawData.reduce((sum, d) => sum + d.usage, 0) / rawData.length : 0;

                if (theoreticalUsage > avgUsage && avgUsage > 0) {
                  return (
                    <div style={{padding: '16px', borderRadius: '8px', background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)', fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '16px'}}>
                      <div style={{fontWeight: '600', color: 'rgb(239, 68, 68)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px'}}>
                        <span style={{fontSize: '1rem'}}>⚠️</span> 입력값 검토 필요
                      </div>
                      <p style={{margin: 0}}>
                        이론 계산된 보일러 사용량 ({Math.round(theoreticalUsage).toLocaleString()} kWh)이 실제 월평균 전체 사용량 ({Math.round(avgUsage).toLocaleString()} kWh)보다 {Math.round(theoreticalUsage / avgUsage * 10) / 10}배 높습니다.
                      </p>
                      <p style={{margin: '8px 0 0 0'}}>
                        💡 <strong>전기용량, 일평균 사용시간, 가동일수를 다시 확인</strong>해주세요. 입력값이 너무 크면 부정확한 결과가 나올 수 있습니다.
                      </p>
                    </div>
                  );
                }
                return null;
              })()}

              <div className="grid-3">
                <div className="form-group">
                  <label>보일러 전기용량</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input
                      type="number"
                      className="form-control"
                      value={boilerCapacity}
                      onChange={e => setBoilerCapacity(e.target.value)}
                      placeholder="0"
                      step="0.1"
                    />
                    <span>kW</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>일평균 사용시간</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input
                      type="number"
                      className="form-control"
                      value={dailyUsageHours}
                      onChange={e => setDailyUsageHours(e.target.value)}
                      placeholder="0"
                      step="0.1"
                    />
                    <span>시간/일</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>년간 가동일</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input
                      type="number"
                      className="form-control"
                      value={annualOperatingDays}
                      onChange={e => setAnnualOperatingDays(e.target.value)}
                      placeholder="365"
                    />
                    <span>일/년</span>
                  </div>
                </div>
              </div>
            </>
          )}
        </section>
      )}

      {/* 2. 시뮬레이션 설정 */}
      {monthlyData.length > 0 && (
        <>
          <div ref={infographicRef}>
            <section className="glass-panel mb-6">
              <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
                <Settings size={24} color="var(--accent-color)" />
                투자비 및 절감율 시뮬레이션
              </h2>
            
            <div style={{background: 'rgba(30, 41, 59, 0.5)', padding: '24px', borderRadius: '12px', marginBottom: '24px', border: '1px solid var(--card-border)'}}>
              <h3 style={{marginBottom: '20px', fontSize: '1.1rem', color: 'var(--text-primary)', display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px'}}>
                <span>상세 설비 투자비 입력</span>
                <span style={{color: 'var(--text-secondary)'}}>설비 합계: {formatMoney(facilityInvestment)}원</span>
              </h3>
              <div className="grid-3">
                <div className="form-group">
                  <label>보일러 대수</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="number" className="form-control" value={boilerCount} onChange={e => setBoilerCount(e.target.value)} min="1" />
                    <span>대</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>보일러 단가</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="text" className="form-control" value={formatInput(boilerUnitPrice)} onChange={handleMoneyChange(setBoilerUnitPrice)} />
                    <span>원</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>대당 설치비</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="text" className="form-control" value={formatInput(installationCost)} onChange={handleMoneyChange(setInstallationCost)} />
                    <span>원</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>축열탱크비 (총액)</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="text" className="form-control" value={formatInput(thermalTankCost)} onChange={handleMoneyChange(setThermalTankCost)} />
                    <span>원</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>전기공사비 (총액)</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="text" className="form-control" value={formatInput(electricalWorkCost)} onChange={handleMoneyChange(setElectricalWorkCost)} />
                    <span>원</span>
                  </div>
                </div>
                <div className="form-group">
                  <label>기타 비용 (총액)</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                    <input type="text" className="form-control" value={formatInput(otherCosts)} onChange={handleMoneyChange(setOtherCosts)} />
                    <span>원</span>
                  </div>
                </div>
              </div>
            </div>

            <h3 style={{marginBottom: '20px', fontSize: '1.1rem', color: 'var(--text-primary)', display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px', borderTop: '1px solid rgba(255,255,255,0.05)', paddingTop: '24px'}}>
              <span>금융 및 보증 비용 (선택)</span>
            </h3>
            <div className="grid-3" style={{marginBottom: '32px'}}>
              <div className="form-group">
                <label>켑코이에스 금융비용 (%)</label>
                <div style={{display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'nowrap'}}>
                  <input type="number" className="form-control" value={kepcoRate} onChange={e => setKepcoRate(e.target.value)} step="0.1" style={{width: '80px'}} />
                  <span style={{whiteSpace: 'nowrap'}}>% <span style={{fontSize: '0.85rem', color: 'var(--text-secondary)'}}>({formatMoney(kepcoCost)}원)</span></span>
                </div>
              </div>
              <div className="form-group">
                <label>SGI서울보증 보증비용 (%)</label>
                <div style={{display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'nowrap'}}>
                  <input type="number" className="form-control" value={sgiRate} onChange={e => setSgiRate(e.target.value)} step="0.1" style={{width: '80px'}} />
                  <span style={{whiteSpace: 'nowrap'}}>% <span style={{fontSize: '0.85rem', color: 'var(--text-secondary)'}}>({formatMoney(sgiCost)}원)</span></span>
                </div>
              </div>
              <div className="form-group" style={{display: 'flex', alignItems: 'flex-end', justifyContent: 'flex-end', marginTop: '12px'}}>
                <div style={{textAlign: 'right', padding: '16px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '8px', border: '1px solid rgba(59, 130, 246, 0.3)', width: '100%'}}>
                  <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '4px'}}>최종 총 비용 (설비 + 금융 + 보증)</div>
                  <div style={{fontSize: '1.5rem', fontWeight: 700, color: 'var(--primary-color)'}}>{formatMoney(totalCost)}원</div>
                </div>
              </div>
            </div>

            <div className="grid-2">
              <div className="form-group">
                <label>할부 / 임대 개월 수</label>
                <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                  <input 
                    type="number" 
                    className="form-control" 
                    value={installmentMonths} 
                    onChange={e => setInstallmentMonths(e.target.value)}
                  />
                  <span>개월</span>
                </div>
              </div>
              <div className="form-group">
                <label>기존 대비 예상 절감율 (%)</label>
                <div style={{display: 'flex', alignItems: 'center', gap: '8px'}}>
                  <input 
                    type="number" 
                    className="form-control" 
                    value={savingsRate} 
                    onChange={e => setSavingsRate(e.target.value)}
                    max="100" min="0"
                  />
                  <span>%</span>
                </div>
                {stats && stats.breakevenRate > 0 && (
                  <div style={{fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '8px'}}>
                    💡 전액 회수(손익분기점) 달성을 위한 최소 <strong>{stats.breakevenRate.toFixed(1)}% 이상</strong> 절감 필요
                  </div>
                )}
              </div>
            </div>
          </section>

          {/* 3. 분석 결과 대시보드 */}
          <section className="mb-6" style={{padding: '24px', background: 'var(--bg-color)', borderRadius: '16px', border: '1px solid var(--card-border)'}}>
            <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
              <TrendingDown size={24} color="var(--success-color)" />
              비교 분석 대시보드
            </h2>
            <div className="grid-3 mb-6">
              <div className="stat-card">
                <Calculator size={32} color="var(--primary-color)" style={{margin: '0 auto'}} />
                <div className="stat-label mt-6">연간 누적 절감 예상액</div>
                <div className="stat-value text-success">{formatMoney(stats.totalSavings)}원</div>
                <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)'}}>
                  (월 평균 {formatMoney(stats.averageSavingsMonthly)}원 절감)
                </div>
              </div>
              <div className="stat-card">
                <DollarSign size={32} color="var(--warning-color)" style={{margin: '0 auto'}} />
                <div className="stat-label mt-6">예상 월 할부/임대료</div>
                <div className="stat-value text-warning">{formatMoney(stats.monthlyInstallment)}원</div>
                <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)'}}>
                  ({installmentMonths}개월 기준)
                </div>
              </div>
              <div className="stat-card" style={{
                background: stats.netBenefitMonthly >= 0 
                  ? 'linear-gradient(145deg, rgba(16,185,129,0.1), rgba(15,23,42,0.9))'
                  : 'linear-gradient(145deg, rgba(239,68,68,0.1), rgba(15,23,42,0.9))',
                borderColor: stats.netBenefitMonthly >= 0 ? 'rgba(16,185,129,0.3)' : 'rgba(239,68,68,0.3)'
              }}>
                <TrendingDown size={32} color={stats.netBenefitMonthly >= 0 ? "var(--success-color)" : "var(--danger-color)"} style={{margin: '0 auto'}} />
                <div className="stat-label mt-6">실질 월평균 순이익 (절감액 - 할부금)</div>
                <div className={`stat-value ${stats.netBenefitMonthly >= 0 ? 'text-success' : 'text-danger'}`}>
                  {stats.netBenefitMonthly > 0 ? '+' : ''}{formatMoney(stats.netBenefitMonthly)}원
                </div>
                <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)'}}>
                  {stats.netBenefitMonthly >= 0 ? '도입 즉시 수익 발생!' : '절감액으로 할부금 일부 상쇄'}
                </div>
              </div>
            </div>

            {/* 개선 후 실제 월 부담 비교 */}
            <div style={{marginBottom: '24px', padding: '24px', borderRadius: '12px', background: 'rgba(30,41,59,0.7)', border: '1px solid var(--card-border)'}}>
              <div style={{fontWeight: 700, fontSize: '1rem', marginBottom: '4px', color: 'var(--text-secondary)', textTransform: 'uppercase', letterSpacing: '0.05em'}}>
                월 평균 부담 비용 비교 (교체 전 vs 교체 후)
              </div>
              <div style={{fontSize: '0.8rem', color: 'rgba(255,255,255,0.35)', marginBottom: '16px'}}>
                * 입력된 12개월 데이터의 월 평균 기준
              </div>
              <div style={{display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap'}}>
                {/* 개선 전 */}
                <div style={{flex: '1 1 180px', padding: '16px', borderRadius: '10px', background: 'rgba(59,130,246,0.1)', border: '1px solid rgba(59,130,246,0.3)', textAlign: 'center'}}>
                  <div style={{fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '6px'}}>교체 전 월 평균 전기료</div>
                  <div style={{fontSize: '1.4rem', fontWeight: 700, color: 'var(--primary-color)'}}>{formatMoney(stats.avgTotalBill)}원</div>
                </div>

                <div style={{fontSize: '1.5rem', color: 'var(--text-secondary)', flexShrink: 0}}>→</div>

                {/* 개선 후 전기료 + 임대료 분해 */}
                <div style={{flex: '1 1 220px', padding: '16px', borderRadius: '10px', background: 'rgba(16,185,129,0.1)', border: '1px solid rgba(16,185,129,0.3)', textAlign: 'center'}}>
                  <div style={{fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '6px'}}>교체 후 실제 월 평균 부담</div>
                  <div style={{fontSize: '1.4rem', fontWeight: 700, color: 'var(--success-color)'}}>{formatMoney(stats.monthlyTotalBurden)}원</div>
                  <div style={{fontSize: '0.78rem', color: 'rgba(255,255,255,0.45)', marginTop: '6px'}}>
                    전기료 평균 {formatMoney(stats.avgNewTotalBill)}원 + 임대료 {formatMoney(stats.monthlyInstallment)}원
                  </div>
                </div>

                <div style={{fontSize: '1.5rem', color: 'var(--text-secondary)', flexShrink: 0}}>=</div>

                {/* 차이 */}
                {(() => {
                  const diff = stats.avgTotalBill - stats.monthlyTotalBurden;
                  return (
                    <div style={{flex: '1 1 160px', padding: '16px', borderRadius: '10px', background: diff >= 0 ? 'rgba(16,185,129,0.15)' : 'rgba(239,68,68,0.1)', border: diff >= 0 ? '1px solid rgba(16,185,129,0.4)' : '1px solid rgba(239,68,68,0.3)', textAlign: 'center'}}>
                      <div style={{fontSize: '0.8rem', color: 'var(--text-secondary)', marginBottom: '6px'}}>{diff >= 0 ? '월 절감' : '월 추가 부담'}</div>
                      <div style={{fontSize: '1.4rem', fontWeight: 700, color: diff >= 0 ? 'var(--success-color)' : 'var(--danger-color)'}}>
                        {diff >= 0 ? '-' : '+'}{formatMoney(Math.abs(diff))}원
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>

            <div className="glass-panel" style={{marginBottom: '24px', background: 'linear-gradient(135deg, rgba(30,41,59,0.8), rgba(15,23,42,0.9))'}}>
              <h3 style={{marginBottom: '20px', color: 'var(--text-primary)', borderBottom: '1px solid var(--card-border)', paddingBottom: '12px', fontSize: '1.25rem', fontWeight: 600}}>장기 투자 수익 분석 (할부 {installmentMonths}개월 기준)</h3>
              <div className="grid-3" style={{gap: '16px'}}>
                <div style={{padding: '16px', background: 'rgba(255,255,255,0.03)', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.05)'}}>
                  <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '4px'}}>할부 기간 내 전체 할부금</div>
                  <div style={{fontSize: '0.75rem', color: 'rgba(255,255,255,0.5)', marginBottom: '8px'}}>기기, 설치비 등 모든 설비 투자비와 금융/보증 등 부대비용을 더하여 계약 기간 동안 고객님이 납부하실 총 비중입니다.</div>
                  <div style={{fontSize: '1.25rem', fontWeight: 600}}>{formatMoney(totalCost)}원</div>
                </div>
                <div style={{padding: '16px', background: 'rgba(255,255,255,0.03)', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.05)'}}>
                  <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '4px'}}>할부 기간 내 전체 절약액</div>
                  <div style={{fontSize: '0.75rem', color: 'rgba(255,255,255,0.5)', marginBottom: '8px'}}>고효율 보일러 가동을 통해 기존 장비와 대비하여 할부/임대 기간 전체에 걸쳐 아낄 수 있는 예상 전기요금 총합입니다.</div>
                  <div style={{fontSize: '1.25rem', fontWeight: 600, color: 'var(--success-color)'}}>+ {formatMoney(stats.averageSavingsMonthly * Number(installmentMonths))}원</div>
                </div>
                <div style={{padding: '16px', background: 'rgba(255,255,255,0.03)', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.05)'}}>
                  <div style={{fontSize: '0.875rem', color: 'var(--text-secondary)', marginBottom: '4px'}}>할부 기간 내 전체 수익액</div>
                  <div style={{fontSize: '0.75rem', color: 'rgba(255,255,255,0.5)', marginBottom: '8px'}}>할부기간 동안 아낀 전기요금 총합에서 각종 총 비용을 뺀 실제 이익금액입니다. 플러스일 경우 갚고도 남는 수익을 뜻합니다.</div>
                  <div style={{fontSize: '1.25rem', fontWeight: 600, color: 'var(--warning-color)'}}>{formatMoney((stats.averageSavingsMonthly * Number(installmentMonths)) - totalCost)}원</div>
                </div>
                <div style={{padding: '24px', background: 'rgba(16,185,129,0.1)', borderRadius: '8px', gridColumn: '1 / -1', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px', border: '1px solid rgba(16,185,129,0.3)'}}>
                  <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                    <div style={{background: 'var(--success-color)', borderRadius: '50%', width: '40px', height: '40px', display: 'flex', alignItems: 'center', justifyContent: 'center'}}>
                      <TrendingDown color="white" size={24} />
                    </div>
                    <div>
                      <div style={{fontSize: '1.25rem', fontWeight: 700, color: 'var(--text-primary)'}}>할부 종료 후 연간 순수익</div>
                      <div style={{fontSize: '0.85rem', color: 'rgba(255,255,255,0.6)', marginTop: '4px'}}>* 의미: 투자비 상환이 모두 끝난 후, 기존 대비 100% 고객님께 돌아오는 매년 누적 예상 순이익입니다.</div>
                    </div>
                  </div>
                  <div style={{fontSize: '2rem', fontWeight: 800, color: 'var(--success-color)'}}>+ {formatMoney(stats.totalSavings)}원 / 년</div>
                </div>
              </div>
            </div>

            <div className="grid-2">
              {/* 차트: 기존요금 vs 개선요금 비교 */}
              <div className="glass-panel">
                <h3>월별 전기요금 비교 (기존 vs 개선 후)</h3>
                <div className="chart-container">
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart data={monthlyData} margin={{top: 20, right: 20, left: 20, bottom: 20}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                      <XAxis dataKey="displayMonth" stroke="var(--text-secondary)" />
                      <YAxis tickFormatter={t => formatMoney(t/10000) + '만'} stroke="var(--text-secondary)" />
                      <RechartsTooltip 
                        formatter={(value) => formatMoney(value) + '원'}
                        labelFormatter={(label) => label}
                        contentStyle={{ backgroundColor: 'var(--bg-color)', border: '1px solid var(--card-border)', borderRadius: '8px' }}
                      />
                      <Legend />
                      <Bar dataKey="totalBill" name="기존 총 요금" fill="rgba(59, 130, 246, 0.4)" radius={[4, 4, 0, 0]} />
                      <Line type="monotone" dataKey="newTotalBill" name="개선 후 총 요금" stroke="var(--success-color)" strokeWidth={3} />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* 차트: 순수 보일러 요금 절감액 */}
              <div className="glass-panel">
                <h3>순수 보일러 요금 및 절감액</h3>
                <div className="chart-container">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={monthlyData} margin={{top: 20, right: 20, left: 20, bottom: 20}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.1)" />
                      <XAxis dataKey="displayMonth" stroke="var(--text-secondary)" />
                      <YAxis tickFormatter={t => formatMoney(t/10000) + '만'} stroke="var(--text-secondary)" />
                      <RechartsTooltip 
                        formatter={(value) => formatMoney(value) + '원'}
                        labelFormatter={(label) => label}
                        contentStyle={{ backgroundColor: 'var(--bg-color)', border: '1px solid var(--card-border)', borderRadius: '8px' }}
                      />
                      <Legend />
                      <Bar dataKey="newBoilerBill" name="고효율 보일러 예상 요금" stackId="a" fill="rgba(16, 185, 129, 0.7)" radius={[0, 0, 4, 4]} />
                      <Bar dataKey="expectedSavings" name="절감액" stackId="a" fill="rgba(239, 68, 68, 0.5)" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          </section>
          </div>

          {/* 4. 데이터 테이블 & 익스포트 */}
          <section className="glass-panel">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px', marginBottom: '20px' }}>
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px', margin: 0 }}>
                <Calendar size={20} color="var(--primary-color)" />
                월별 상세 데이터 리포트 목록
              </h3>
              <div style={{display: 'flex', gap: '12px', flexWrap: 'wrap'}}>
                <button className="btn" style={{background: 'rgba(59, 130, 246, 0.2)', color: 'var(--primary-color)'}} onClick={handleDownloadImage}>
                  <Download size={18} /> 대시보드 캡처 (이미지)
                </button>
                <button className="btn btn-accent" onClick={handleDownloadExcel}>
                  <Download size={18} /> 보고서 다운로드 (Excel)
                </button>
              </div>
            </div>
            <div className="table-wrapper">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>월</th>
                    <th>총 전력량(kWh)</th>
                    <th>기존 총 요금</th>
                    <th>기존 보일러 요금</th>
                    <th>개선 후 보일러 요금</th>
                    <th className="text-success">예상 절감액</th>
                  </tr>
                </thead>
                <tbody>
                  {monthlyData.map(d => (
                    <tr key={d.displayMonth}>
                      <td style={{textAlign: 'center'}}>{d.displayMonth}</td>
                      <td>{d.rawUsage !== null ? formatMoney(d.usage) : '-'}</td>
                      <td>{formatMoney(d.totalBill)}원</td>
                      <td>{formatMoney(d.boilerBill)}원</td>
                      <td>{formatMoney(d.newBoilerBill)}원</td>
                      <td className="text-success" style={{fontWeight: 600}}>↓ {formatMoney(d.expectedSavings)}원</td>
                    </tr>
                  ))}
                  <tr style={{background: 'rgba(59, 130, 246, 0.1)', fontWeight: 'bold'}}>
                    <td style={{textAlign: 'center'}}>합계</td>
                    <td>{stats.usageAvailable ? formatMoney(stats.totalUsage) : '-'}</td>
                    <td>{formatMoney(stats.totalBill)}원</td>
                    <td>{formatMoney(stats.totalBoilerBill)}원</td>
                    <td>{formatMoney(stats.totalBoilerBill - stats.totalSavings)}원</td>
                    <td className="text-success">↓ {formatMoney(stats.totalSavings)}원</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </section>
        </>
      )}
    </div>
  );
}
