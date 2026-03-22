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
  const [otherPowerBill, setOtherPowerBill] = useState(''); // 월 비보일러 전기요금
  
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

  const fileInputRef = useRef(null);
  const infographicRef = useRef(null);

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

    // Estimate usage based on average electricity rate (약 120원/kWh 기준)
    const estimatedTotalUsage = Math.round(totalBill / 120);

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

  // Parse Excel
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      let parsed = [];
      // 간단한 파서: 숫자 데이터를 찾아 월별로 매핑 (첫번째 시도의 첫 3개 숫자/날짜성 컬럼을 사용)
      // 혹은 헤더를 기반으로 탐색
      let headerRowIndex = -1;
      let monthIndex = -1, usageIndex = -1, billIndex = -1;

      for (let i = 0; i < Math.min(20, data.length); i++) {
        const row = data[i];
        if (!row) continue;
        const rowStr = row.join(' ').toLowerCase();
        if ((rowStr.includes('월') || rowStr.includes('년') || rowStr.includes('date')) &&
            (rowStr.includes('사용') || rowStr.includes('kwh') || rowStr.includes('전력')) &&
            (rowStr.includes('요금') || rowStr.includes('금액') || rowStr.includes('청구'))) {
          headerRowIndex = i;
          row.forEach((col, idx) => {
            if (!col) return;
            const cStr = String(col).toLowerCase().replace(/\s/g, '');
            if (!cStr.includes('기간') && (cStr.includes('사용') || cStr.includes('kwh') || cStr.includes('전력'))) {
              usageIndex = usageIndex === -1 ? idx : usageIndex;
            } else if (cStr.includes('요금') || cStr.includes('금액') || cStr.includes('청구')) {
              billIndex = billIndex === -1 ? idx : billIndex;
            } else if (!cStr.includes('기간') && (cStr.includes('월') || cStr.includes('년') || cStr.includes('일') || cStr.includes('date'))) {
              monthIndex = monthIndex === -1 ? idx : monthIndex;
            }
          });
          break;
        }
      }

      // 만약 헤더를 못찾았다면, 대략 첫번째 문자열 컬럼은 날짜, 두번째, 세번째 숫자 컬럼은 사용량/요금으로 추정
      if (headerRowIndex === -1) {
        monthIndex = 0; usageIndex = 1; billIndex = 2;
        headerRowIndex = 0; // 그냥 1번째 줄부터 데이터라고 가정함 (숫자 파싱때 걸러짐)
      } else if (monthIndex === -1 || usageIndex === -1 || billIndex === -1) {
        monthIndex = 0; usageIndex = 1; billIndex = 2;
      }

      for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length < 3) continue;
        
        // 데이터 파싱
        let rawDate = row[monthIndex];
        let usage = parseFloat(String(row[usageIndex]).replace(/,/g, ''));
        let bill = parseFloat(String(row[billIndex]).replace(/,/g, ''));

        if (isNaN(usage) || isNaN(bill)) continue;

        let year = null;
        let month = null;

        if (typeof rawDate === 'number' && rawDate > 10000) {
          const date = new Date((rawDate - (25567 + 2)) * 86400 * 1000); // adjust for excel leap year bug & origin
          year = date.getFullYear();
          month = date.getMonth() + 1;
        } else if (rawDate) {
          const dStr = String(rawDate).trim();
          let match = dStr.match(/^(\d{4})[-\.년\s]*(\d{1,2})/);
          if (match) {
            year = parseInt(match[1], 10);
            month = parseInt(match[2], 10);
          } else {
            match = dStr.match(/^(\d{2})[-\.년\s]*(\d{1,2})/);
            if (match) {
              year = 2000 + parseInt(match[1], 10);
              month = parseInt(match[2], 10);
            }
          }
        }

        if (year && month) {
          parsed.push({ year, month, usage, bill });
        }
      }

      // 만약 정규 파싱 실패시, 모의 데이터 제공 기능 필요?
      if (parsed.length === 0) {
        alert("데이터를 헤더(년월/사용량/청구금액)에 맞게 인식하지 못했습니다. 샘플 데이터를 띄웁니다.");
        const currentYear = new Date().getFullYear();
        parsed = Array.from({length: 12}, (_, i) => ({
          year: currentYear,
          month: i + 1,
          usage: Math.floor(10000 + Math.random() * 5000),
          bill: Math.floor(1500000 + Math.random() * 500000)
        }));
      }

      setRawData(parsed);

      const uniqueYears = [...new Set(parsed.map(item => item.year))].sort((a,b)=>b-a);
      setYears(uniqueYears);
      if (uniqueYears.length > 0) {
        setStartYear(uniqueYears[0]);
        setStartMonth(1);
      }
    };
    reader.readAsBinaryString(file);
  };

  // Generate Monthly Template
  const monthlyData = useMemo(() => {
    if (!startYear || !startMonth || rawData.length === 0) return [];
    
    let res = [];
    let currentY = Number(startYear);
    let currentM = Number(startMonth);

    for (let c = 0; c < 12; c++) {
      const found = rawData.find(d => d.year === currentY && d.month === currentM);
      const usage = found ? found.usage : 0;
      const totalBill = found ? found.bill : 0;
      
      const safeOtherPowerBill = Number(otherPowerBill) || 0;
      const safeSavingsRate = Number(savingsRate) || 0;

      const boilerBill = Math.max(0, totalBill - safeOtherPowerBill);
      const savings = boilerBill * (safeSavingsRate / 100);
      const newBoilerBill = boilerBill - savings;

      res.push({
        displayMonth: `${currentY.toString().slice(-2)}년 ${currentM}월`,
        year: currentY,
        month: currentM,
        usage,
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
  }, [rawData, startYear, startMonth, otherPowerBill, savingsRate]);

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
    let totalUsage = 0, totalBill = 0, totalBoilerBill = 0, totalSavings = 0;
    
    monthlyData.forEach(d => {
      totalUsage += d.usage;
      totalBill += d.totalBill;
      totalBoilerBill += d.boilerBill;
      totalSavings += d.expectedSavings;
    });

    const monthlyInstallment = (Number(installmentMonths)||0) > 0 ? totalCost / Number(installmentMonths) : 0;
    // 임대료를 할부금과 동일하게 보거나, 부대비용 포함 가능. 여기서는 할부금=임대료 기준으로 안내.
    const monthlyRental = monthlyInstallment; 

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
                <div className="form-group">
                  <label>월별 비보일러 전기요금 (기본 요금 등)</label>
                  <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
                    <input
                      type="text"
                      className="form-control"
                      value={formatInput(otherPowerBill)}
                      onChange={handleMoneyChange(setOtherPowerBill)}
                    />
                    <span>원/월</span>
                  </div>
                  <div style={{fontSize: '0.8rem', color: 'var(--text-secondary)', marginTop: '8px'}}>
                    이 금액을 뺀 나머지 금액이 "순수 보일러 전기요금"으로 산정됩니다.
                  </div>
                </div>
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
                        ⚡ 사용량은 평균 전기요금 단가(120원/kWh)를 기준으로 자동 추정됩니다
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

      {/* 비보일러 전기요금 입력 (수동 입력 모드일 때만 별도 섹션으로) */}
      {inputMode === 'manual' && monthlyData.length > 0 && (
        <section className="glass-panel mb-6">
          <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
            <Settings size={20} color="var(--primary-color)" />
            추가 설정
          </h2>
          <div className="form-group">
            <label>월별 비보일러 전기요금 (기본 요금 등)</label>
            <div style={{display: 'flex', alignItems: 'center', gap: '12px'}}>
              <input
                type="text"
                className="form-control"
                value={formatInput(otherPowerBill)}
                onChange={handleMoneyChange(setOtherPowerBill)}
              />
              <span>원/월</span>
            </div>
            <div style={{fontSize: '0.8rem', color: 'var(--text-secondary)', marginTop: '8px'}}>
              이 금액을 뺀 나머지 금액이 "순수 보일러 전기요금"으로 산정됩니다.
            </div>
          </div>
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
                <div style={{padding: '24px', background: 'rgba(16,185,129,0.1)', borderRadius: '8px', gridColumn: '1 / -1', display: 'flex', justifyContent: 'space-between', alignItems: 'center', border: '1px solid rgba(16,185,129,0.3)'}}>
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
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <h3 style={{ display: 'flex', alignItems: 'center', gap: '8px', margin: 0 }}>
                <Calendar size={20} color="var(--primary-color)" />
                월별 상세 데이터 리포트 목록
              </h3>
              <div style={{display: 'flex', gap: '12px'}}>
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
                      <td>{formatMoney(d.usage)}</td>
                      <td>{formatMoney(d.totalBill)}원</td>
                      <td>{formatMoney(d.boilerBill)}원</td>
                      <td>{formatMoney(d.newBoilerBill)}원</td>
                      <td className="text-success" style={{fontWeight: 600}}>↓ {formatMoney(d.expectedSavings)}원</td>
                    </tr>
                  ))}
                  <tr style={{background: 'rgba(59, 130, 246, 0.1)', fontWeight: 'bold'}}>
                    <td style={{textAlign: 'center'}}>합계</td>
                    <td>{formatMoney(stats.totalUsage)}</td>
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
