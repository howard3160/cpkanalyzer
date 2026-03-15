'use client';
/* eslint-disable @typescript-eslint/no-explicit-any */

import { useState, useEffect, useRef } from 'react';
import Script from 'next/script';

export default function Home() {
  const [currentLanguage, setCurrentLanguage] = useState('zh-TW');
  const [measurements, setMeasurements] = useState<number[]>([]);
  const [usl, setUsl] = useState<number | null>(null);
  const [lsl, setLsl] = useState<number | null>(null);
  const [stats, setStats] = useState<any>(null);
  const [errorMsg, setErrorMsg] = useState('');
  const [warningMsg, setWarningMsg] = useState('');
  const chartRefs = useRef<Record<string, any>>({});

  // Internationalization
  const i18n = {
    'zh-TW': {
      title: 'CPK 分析',
      dataInputTitle: '資料輸入',
      runAnalysisButton: '▶ 執行分析',
      resultsTitle: '分析結果',
      exportResultsButton: '匯出結果 (CSV)',
      n: '樣本數',
      mean: '平均值',
      stdDev: '標準差',
      min: '最小值',
      max: '最大值',
      cp: 'Cp',
      cpk: 'Cpk',
      pp: 'Pp',
      ppk: 'Ppk',
      sigmaLevel: 'Sigma 等級',
      ppm: '預估 PPM',
      specificationTitle: '製程規格設定',
      uslLabel: '上規格限 (USL)',
      lslLabel: '下規格限 (LSL)',
      errorMissingData: '請先輸入資料',
      errorMissingUSL: '請輸入上規格限 (USL)',
      errorMissingLSL: '請輸入下規格限 (LSL)',
      errorInvalidSpecs: '規格限無效：LSL 應小於 USL',
      errorInsufficientData: '需要至少 2 筆資料',
    },
    en: {
      title: 'CPK Analysis',
      dataInputTitle: 'Data Input',
      runAnalysisButton: '▶ Run Analysis',
      resultsTitle: 'Analysis Results',
      exportResultsButton: 'Export Results (CSV)',
      n: 'Sample Count',
      mean: 'Mean',
      stdDev: 'Std Dev',
      min: 'Minimum',
      max: 'Maximum',
      cp: 'Cp',
      cpk: 'Cpk',
      pp: 'Pp',
      ppk: 'Ppk',
      sigmaLevel: 'Sigma Level',
      ppm: 'Est. PPM',
      specificationTitle: 'Process Specification',
      uslLabel: 'Upper Spec Limit (USL)',
      lslLabel: 'Lower Spec Limit (LSL)',
      errorMissingData: 'Please enter data first',
      errorMissingUSL: 'Please enter USL',
      errorMissingLSL: 'Please enter LSL',
      errorInvalidSpecs: 'Invalid specs: LSL must be less than USL',
      errorInsufficientData: 'At least 2 data points required',
    },
    ja: {
      title: 'CPK 分析',
      dataInputTitle: 'データ入力',
      runAnalysisButton: '▶ 分析を実行',
      resultsTitle: '分析結果',
      exportResultsButton: '結果をエクスポート (CSV)',
      n: 'サンプル数',
      mean: '平均',
      stdDev: '標準偏差',
      min: '最小値',
      max: '最大値',
      cp: 'Cp',
      cpk: 'Cpk',
      pp: 'Pp',
      ppk: 'Ppk',
      sigmaLevel: 'シグマレベル',
      ppm: '推定 PPM',
      specificationTitle: 'プロセス仕様',
      uslLabel: '上限規格 (USL)',
      lslLabel: '下限規格 (LSL)',
      errorMissingData: 'データを入力してください',
      errorMissingUSL: 'USL を入力してください',
      errorMissingLSL: 'LSL を入力してください',
      errorInvalidSpecs: '無効な仕様: LSL は USL より小さい必要があります',
      errorInsufficientData: '最低 2 つのデータポイントが必要',
    },
  };

  const t = (key: string) => {
    const langKey = currentLanguage as keyof typeof i18n;
    const lang = i18n[langKey];
    return lang[key as keyof typeof lang] || key;
  };

  // 格式化數值：最多顯示一個小數位，但去除不必要的0
  const formatNumber = (num: number, decimals: number = 1): string => {
    if (num === null || num === undefined) return 'N/A';
    const rounded = parseFloat(num.toFixed(decimals));
    const str = rounded.toString();
    // 如果沒有小數點，保持原樣；否則去除尾部的0
    if (!str.includes('.')) return str;
    return str.replace(/\.?0+$/, '');
  };

  // Statistical functions
  const erf = (x: number) => {
    const a1 = 0.254829592;
    const a2 = -0.284496736;
    const a3 = 1.421413741;
    const a4 = -1.453152027;
    const a5 = 1.061405429;
    const p = 0.3275911;
    const sign = x < 0 ? -1 : 1;
    x = Math.abs(x);
    const t = 1.0 / (1.0 + p * x);
    const y = 1.0 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
    return sign * y;
  };

  const normalCDF = (x: number, mu: number, sigma: number) => {
    return 0.5 * (1 + erf((x - mu) / (sigma * Math.sqrt(2))));
  };

  const normalPDF = (x: number, mu: number, sigma: number) => {
    return (1 / (sigma * Math.sqrt(2 * Math.PI))) * Math.exp(-0.5 * Math.pow((x - mu) / sigma, 2));
  };

  const inverseNormalCDFApprox = (p: number): number => {
    const a1 = -3.969683028665376e1;
    const a2 = 2.221222899801429e2;
    const a3 = -2.821152023902548e2;
    const a4 = 1.340426573961423e2;
    const a5 = -1.627296189588725;
    const b1 = -5.447609879822406e1;
    const b2 = 1.615858368580409e2;
    const b3 = -1.556989798598866e2;
    const b4 = 6.680131188771972e1;
    const b5 = -7.028495245500852e-1;

    const q = p < 0.02425;
    const r = q ? Math.sqrt(-2.0 * Math.log(p)) : Math.sqrt(-2.0 * Math.log(0.5 - Math.abs(p - 0.5)));
    const u = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) / (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)) * r;
    return q ? -u : p < 0.5 ? -u : u;
  };

  const inverseNormalCDF = (p: number): number => {
    if (p < 0.5) {
      return -inverseNormalCDFApprox(p);
    } else {
      return inverseNormalCDFApprox(1 - p);
    }
  };

  const calculateStats = (data: number[]) => {
    if (data.length < 2 || !usl || !lsl) return null;
    const n = data.length;
    const mean = data.reduce((a, b) => a + b, 0) / n;
    const variance = data.reduce((sum, x) => sum + Math.pow(x - mean, 2), 0) / (n - 1);
    const sigma = Math.sqrt(variance);

    if (sigma === 0) {
      return { n, mean, variance, sigma, min: Math.min(...data), max: Math.max(...data) };
    }

    const uslVal = usl;
    const lslVal = lsl;
    const cpu = (uslVal - mean) / (3 * sigma);
    const cpl = (mean - lslVal) / (3 * sigma);
    const cpk = Math.min(cpu, cpl);
    const cp = (uslVal - lslVal) / (6 * sigma);
    const pp = cp;
    const ppk = cpk;
    const sigmaLevel = cpk * 3;
    const ppmUpper = (1 - normalCDF(uslVal, mean, sigma)) * 1000000;
    const ppmLower = normalCDF(lslVal, mean, sigma) * 1000000;
    const ppm = ppmUpper + ppmLower;

    return {
      n, mean, variance, sigma,
      min: Math.min(...data),
      max: Math.max(...data),
      cp, cpu, cpl, cpk,
      pp, ppk,
      sigmaLevel, ppm,
    };
  };

  // 圖表渲染 - useEffect
  useEffect(() => {
    if (!stats) return;

    const Chart = (window as any).Chart;
    if (!Chart) return;

    const mean = stats.mean;
    const sigma = stats.sigma;
    const data = measurements;

    // 清理舊圖表
    Object.values(chartRefs.current).forEach((chart: any) => {
      if (chart && typeof chart.destroy === 'function') {
        chart.destroy();
      }
    });
    chartRefs.current = {};

    // 1. Normal Distribution
    const normalCanvas = document.getElementById('normalDistChart') as HTMLCanvasElement;
    if (normalCanvas) {
      const xMin = lsl! - 3 * sigma;
      const xMax = usl! + 3 * sigma;
      const step = (xMax - xMin) / 200;
      const labels: string[] = [];
      const pdfData: number[] = [];

      for (let x = xMin; x <= xMax; x += step) {
        labels.push(x.toFixed(1));
        pdfData.push(normalPDF(x, mean, sigma) * 100);
      }

      const ctx = normalCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['normalDistChart'] = new Chart(ctx, {
          type: 'line',
          data: {
            labels,
            datasets: [{
              label: 'Normal Distribution',
              data: pdfData,
              borderColor: '#2563EB',
              backgroundColor: 'rgba(37, 99, 235, 0.1)',
              tension: 0.4,
              fill: true,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { y: { beginAtZero: true } },
          },
        });
      }
    }

    // 2. Histogram
    const histCanvas = document.getElementById('histogramChart') as HTMLCanvasElement;
    if (histCanvas) {
      const k = Math.ceil(Math.log2(data.length) + 1);
      const min = Math.min(...data);
      const max = Math.max(...data);
      const binWidth = (max - min) / k;

      const bins = Array(k).fill(0);
      const binLabels: string[] = [];

      for (let i = 0; i < k; i++) {
        binLabels.push((min + i * binWidth).toFixed(1));
      }

      data.forEach((value) => {
        const binIndex = Math.min(Math.floor((value - min) / binWidth), k - 1);
        bins[binIndex]++;
      });

      const ctx = histCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['histogramChart'] = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: binLabels,
            datasets: [{
              label: 'Frequency',
              data: bins,
              backgroundColor: '#2563EB',
              borderColor: '#1D4ED8',
              borderWidth: 1,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { y: { beginAtZero: true } },
          },
        });
      }
    }

    // 3. X-bar Control Chart
    const xbarCanvas = document.getElementById('xbarChart') as HTMLCanvasElement;
    if (xbarCanvas) {
      const ucl = mean + 3 * sigma;
      const lcl = Math.max(0, mean - 3 * sigma);
      const labels = data.map((_, i) => (i + 1).toString());

      const ctx = xbarCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['xbarChart'] = new Chart(ctx, {
          type: 'line',
          data: {
            labels,
            datasets: [
              {
                label: 'Data Points',
                data,
                borderColor: '#2563EB',
                backgroundColor: 'rgba(37, 99, 235, 0.1)',
                pointRadius: 4,
                borderWidth: 2,
              },
              {
                label: 'UCL',
                data: Array(data.length).fill(ucl),
                borderColor: '#DC2626',
                borderDash: [5, 5],
                fill: false,
                pointRadius: 0,
              },
              {
                label: 'CL',
                data: Array(data.length).fill(mean),
                borderColor: '#059669',
                borderDash: [5, 5],
                fill: false,
                pointRadius: 0,
              },
              {
                label: 'LCL',
                data: Array(data.length).fill(lcl),
                borderColor: '#DC2626',
                borderDash: [5, 5],
                fill: false,
                pointRadius: 0,
              },
            ],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: true } },
            scales: { y: { beginAtZero: false } },
          },
        });
      }
    }

    // 4. Box Plot
    const boxCanvas = document.getElementById('boxPlotChart') as HTMLCanvasElement;
    if (boxCanvas) {
      const sorted = [...data].sort((a, b) => a - b);
      const q1 = sorted[Math.floor(sorted.length * 0.25)];
      const median = sorted[Math.floor(sorted.length * 0.5)];
      const q3 = sorted[Math.floor(sorted.length * 0.75)];
      const min = sorted[0];
      const max = sorted[sorted.length - 1];
      const iqr = q3 - q1;

      const ctx = boxCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['boxPlotChart'] = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: ['Data Distribution'],
            datasets: [
              {
                label: 'Whiskers (Min-Max)',
                data: [max - min],
                backgroundColor: 'rgba(200, 200, 200, 0.3)',
                borderColor: 'rgba(100, 100, 100, 0.5)',
                borderWidth: 1,
              },
              {
                label: 'IQR (Q1-Q3)',
                data: [iqr],
                backgroundColor: 'rgba(37, 99, 235, 0.5)',
                borderColor: '#2563EB',
                borderWidth: 2,
              },
            ],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
              legend: { display: true },
              tooltip: {
                callbacks: {
                  afterLabel: () => `Q1: ${q1.toFixed(2)} | Median: ${median.toFixed(2)} | Q3: ${q3.toFixed(2)}`
                }
              }
            },
            scales: { y: { beginAtZero: true } },
          },
        });
      }
    }

    // 5. Trend Line
    const trendCanvas = document.getElementById('trendChart') as HTMLCanvasElement;
    if (trendCanvas) {
      const labels = data.map((_, i) => (i + 1).toString());
      const windowSize = 5;
      const movingAvg: number[] = [];

      for (let i = 0; i < data.length; i++) {
        const start = Math.max(0, i - Math.floor(windowSize / 2));
        const end = Math.min(data.length, i + Math.floor(windowSize / 2) + 1);
        const avg = data.slice(start, end).reduce((a, b) => a + b, 0) / (end - start);
        movingAvg.push(avg);
      }

      const ctx = trendCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['trendChart'] = new Chart(ctx, {
          type: 'line',
          data: {
            labels,
            datasets: [
              {
                label: 'Data Points',
                data,
                borderColor: '#2563EB',
                backgroundColor: 'rgba(37, 99, 235, 0.1)',
                borderWidth: 1,
                pointRadius: 3,
              },
              {
                label: 'Moving Avg',
                data: movingAvg,
                borderColor: '#D97706',
                borderWidth: 2,
                fill: false,
                pointRadius: 0,
              },
              {
                label: 'USL',
                data: Array(data.length).fill(usl!),
                borderColor: '#DC2626',
                borderDash: [5, 5],
                fill: false,
                pointRadius: 0,
              },
              {
                label: 'LSL',
                data: Array(data.length).fill(lsl!),
                borderColor: '#DC2626',
                borderDash: [5, 5],
                fill: false,
                pointRadius: 0,
              },
            ],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: true } },
          },
        });
      }
    }

    // 6. CPK Gauge
    const gaugeCanvas = document.getElementById('cpkGaugeChart') as HTMLCanvasElement;
    if (gaugeCanvas && stats.cpk !== undefined) {
      let color = '#DC2626';
      if (stats.cpk >= 1.67) color = '#16A34A';
      else if (stats.cpk >= 1.33) color = '#2563EB';
      else if (stats.cpk >= 1.0) color = '#D97706';

      const value = Math.min(stats.cpk, 2);

      const ctx = gaugeCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['cpkGaugeChart'] = new Chart(ctx, {
          type: 'doughnut',
          data: {
            labels: ['Cpk', 'Remaining'],
            datasets: [{
              data: [value, 2 - value],
              backgroundColor: [color, '#E5E7EB'],
              borderWidth: 0,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false }, tooltip: { enabled: false } },
          },
        });
      }
    }

    // 7. Yield & PPM
    const yieldCanvas = document.getElementById('yieldChart') as HTMLCanvasElement;
    if (yieldCanvas && stats.ppm !== undefined) {
      const yield_ = Math.max(0, 100 - (stats.ppm / 10000));

      const ctx = yieldCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['yieldChart'] = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: ['Process Yield'],
            datasets: [
              {
                label: 'Yield %',
                data: [yield_],
                backgroundColor: yield_ > 95 ? '#16A34A' : yield_ > 80 ? '#D97706' : '#DC2626',
                borderColor: '#1f2937',
                borderWidth: 1,
              },
            ],
          },
          options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
              legend: { display: false },
              tooltip: {
                callbacks: {
                  label: () => `Yield: ${yield_.toFixed(2)}% | PPM: ${stats.ppm.toFixed(0)}`
                }
              }
            },
            scales: {
              x: { min: 0, max: 100 }
            }
          },
        });
      }
    }

    // 8. Q-Q Plot
    const qqCanvas = document.getElementById('qqPlotChart') as HTMLCanvasElement;
    if (qqCanvas) {
      const sorted = [...data].sort((a, b) => a - b);
      const n = sorted.length;
      const theoreticalQuantiles: number[] = [];

      for (let i = 0; i < n; i++) {
        const p = (i + 1) / (n + 1);
        const z = inverseNormalCDF(p);
        theoreticalQuantiles.push(mean + z * sigma);
      }

      const ctx = qqCanvas.getContext('2d');
      if (ctx) {
        chartRefs.current['qqPlotChart'] = new Chart(ctx, {
          type: 'scatter',
          data: {
            datasets: [{
              label: 'Q-Q Plot',
              data: sorted.map((y, i) => ({
                x: theoreticalQuantiles[i],
                y: y,
              })),
              backgroundColor: '#2563EB',
              pointRadius: 4,
            }],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
              x: { type: 'linear', position: 'bottom' },
              y: { type: 'linear' },
            },
          },
        });
      }
    }
  }, [stats, measurements, usl, lsl]); // eslint-disable-line react-hooks/exhaustive-deps

  const runAnalysis = () => {
    setErrorMsg('');
    setWarningMsg('');

    if (measurements.length === 0) {
      setErrorMsg(t('errorMissingData'));
      return;
    }

    if (usl === null || isNaN(usl)) {
      setErrorMsg(t('errorMissingUSL'));
      return;
    }

    if (lsl === null || isNaN(lsl)) {
      setErrorMsg(t('errorMissingLSL'));
      return;
    }

    if (lsl >= usl) {
      setErrorMsg(t('errorInvalidSpecs'));
      return;
    }

    if (measurements.length < 2) {
      setErrorMsg(t('errorInsufficientData'));
      return;
    }

    const calculatedStats = calculateStats(measurements);
    if (!calculatedStats) {
      setErrorMsg(t('errorInsufficientData'));
      return;
    }

    if (calculatedStats.sigma === 0) {
      setErrorMsg('標準差為零，無法計算 Cpk');
      return;
    }

    setStats(calculatedStats);
  };

  const exportResults = () => {
    if (!stats) {
      setWarningMsg('請先執行分析');
      return;
    }

    const csv = [
      ['CPK 分析結果 / CPK Analysis Results'],
      [],
      ['項目 / Item', '數值 / Value', '單位 / Unit'],
      ['樣本數 / Sample Count', stats.n, ''],
      ['平均值 / Mean', stats.mean.toFixed(4), ''],
      ['標準差 / Std Dev', stats.sigma.toFixed(4), ''],
      ['最小值 / Minimum', stats.min.toFixed(4), ''],
      ['最大值 / Maximum', stats.max.toFixed(4), ''],
      ['上規格限 / USL', usl, ''],
      ['下規格限 / LSL', lsl, ''],
      ['Cp', stats.cp?.toFixed(4) || 'N/A', ''],
      ['Cpk', stats.cpk?.toFixed(4) || 'N/A', ''],
      ['Pp', stats.pp?.toFixed(4) || 'N/A', ''],
      ['Ppk', stats.ppk?.toFixed(4) || 'N/A', ''],
      ['Sigma Level', stats.sigmaLevel?.toFixed(2) || 'N/A', 'σ'],
      ['預估 PPM / Est. PPM', stats.ppm?.toFixed(0) || 'N/A', ''],
      [],
      ['原始資料 / Raw Data'],
      ...measurements.map((m) => [m]),
    ];

    const csvContent = csv.map((row) => row.join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'cpk_analysis_results.csv';
    link.click();
  };

  return (
    <>
      <Script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js" />
      <Script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js" />

      <style jsx global>{`
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background-color: #f4f4f4; color: #333; }
        .header { background: #ffffff; padding: 20px 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
        .header h1 { font-size: 28px; font-weight: 700; color: #1f2937; }
        .language-switcher { display: flex; gap: 8px; }
        .lang-btn { padding: 8px 12px; border: 1px solid #e5e7eb; background: #ffffff; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; transition: all 0.2s; }
        .lang-btn.active { background: #2563eb; color: #ffffff; border-color: #2563eb; }
        .lang-btn:hover { border-color: #2563eb; }
        .container { max-width: 1400px; margin: 0 auto; padding: 24px; }
        .card { background: #ffffff; border-radius: 16px; box-shadow: 0 4px 24px rgba(0,0,0,0.07); padding: 24px; margin-bottom: 24px; }
        .card h2 { font-size: 18px; font-weight: 600; margin-bottom: 16px; color: #1f2937; }
        .spec-inputs { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 16px; }
        .spec-input-group { display: flex; flex-direction: column; }
        .spec-input-group label { font-size: 13px; font-weight: 600; color: #374151; margin-bottom: 6px; }
        .spec-input-group input { padding: 10px 12px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; }
        .btn-primary { background: #2563eb; color: #ffffff; width: 100%; padding: 16px; font-size: 16px; border: none; border-radius: 8px; font-weight: 600; cursor: pointer; margin-top: 16px; }
        .btn-primary:hover { background: #1d4ed8; }
        .error-message { background: #fee2e2; color: #dc2626; padding: 12px; border-radius: 6px; margin-bottom: 12px; font-size: 13px; }
        .warning-message { background: #fef3c7; color: #d97706; padding: 12px; border-radius: 6px; margin-bottom: 12px; font-size: 13px; }
        .results-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .result-item { padding: 16px; background: #f9fafb; border-radius: 8px; text-align: center; border-left: 4px solid #d1d5db; }
        .result-item.excellent { background: #ecfdf5; border-left-color: #16a34a; }
        .result-item.good { background: #eff6ff; border-left-color: #2563eb; }
        .result-item.marginal { background: #fffbeb; border-left-color: #d97706; }
        .result-item.poor { background: #fee2e2; border-left-color: #dc2626; }
        .result-label { font-size: 11px; font-weight: 600; color: #6b7280; text-transform: uppercase; margin-bottom: 8px; }
        .result-value { font-size: 22px; font-weight: 700; color: #1f2937; }
        .result-item.excellent .result-value { color: #16a34a; }
        .result-item.good .result-value { color: #2563eb; }
        .result-item.marginal .result-value { color: #d97706; }
        .result-item.poor .result-value { color: #dc2626; }
        .btn-export { background: #10b981; color: #ffffff; padding: 12px 16px; border: none; border-radius: 8px; cursor: pointer; font-weight: 600; transition: all 0.2s; margin-top: 16px; }
        .btn-export:hover { background: #059669; }
        .charts-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(500px, 1fr)); gap: 24px; margin-bottom: 24px; }
        .chart-card { background: #ffffff; border-radius: 16px; box-shadow: 0 4px 24px rgba(0,0,0,0.07); padding: 24px; min-height: 400px; }
        .chart-card h3 { font-size: 15px; font-weight: 600; margin-bottom: 16px; color: #1f2937; }
        .chart-container { width: 100%; height: 320px; position: relative; }
        @media (max-width: 1024px) { .charts-grid { grid-template-columns: 1fr; } }
      `}</style>

      <div className="header">
        <h1>{t('title')}</h1>
        <div className="language-switcher">
          <select 
            value={currentLanguage} 
            onChange={(e) => setCurrentLanguage(e.target.value)}
            style={{
              padding: '8px 12px',
              border: '1px solid #e5e7eb',
              background: '#ffffff',
              borderRadius: '6px',
              cursor: 'pointer',
              fontSize: '14px',
              fontWeight: '500',
            }}
          >
            <option value="zh-TW">繁體中文</option>
            <option value="en">English</option>
            <option value="ja">日本語</option>
          </select>
        </div>
      </div>

      <div className="container">
        {errorMsg && <div className="error-message">{errorMsg}</div>}
        {warningMsg && <div className="warning-message">{warningMsg}</div>}

        <div className="card">
          <h2>{t('dataInputTitle')}</h2>
          <div style={{ marginBottom: '16px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500' }}>
              輸入測量資料 (支援空格分隔、換行分隔或上傳 Excel):
            </label>
            <textarea
              style={{
                width: '100%',
                height: '150px',
                padding: '10px',
                border: '1px solid #d1d5db',
                borderRadius: '6px',
                fontSize: '14px',
                fontFamily: 'monospace',
              }}
              placeholder="例如：90 91 92&#10;或：90&#10;91&#10;92"
              onChange={(e) => {
                const text = e.target.value;
                // 支援空格或換行分隔
                const values = text
                  .split(/[\s,\n]+/)
                  .map((v) => parseFloat(v.trim()))
                  .filter((v) => !isNaN(v));
                setMeasurements(values);
              }}
            ></textarea>
            <small style={{ color: '#6b7280', marginTop: '8px', display: 'block' }}>
              已輸入 {measurements.length} 筆資料
            </small>
          </div>

          <div style={{ borderTop: '1px solid #e5e7eb', paddingTop: '16px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontWeight: '500' }}>
              或上傳 Excel 文件：
            </label>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = (event) => {
                  try {
                    const data = new Uint8Array(event.target?.result as ArrayBuffer);
                    const XLSX = (window as any).XLSX;
                    if (!XLSX) {
                      alert('Please wait for libraries to load');
                      return;
                    }
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // 提取第一列的數值
                    const values = (jsonData as any[])
                      .slice(1) // 跳過標題
                      .map((row: any) => {
                        const val = Array.isArray(row) ? row[0] : row;
                        return parseFloat(val);
                      })
                      .filter((v) => !isNaN(v));
                    
                    setMeasurements(values);
                  } catch (error) {
                    alert('Error reading file: ' + (error as Error).message);
                  }
                };
                reader.readAsArrayBuffer(file);
              }}
              style={{
                padding: '8px 12px',
                border: '1px solid #d1d5db',
                borderRadius: '6px',
                cursor: 'pointer',
              }}
            />
          </div>
        </div>

        <div className="card">
          <h2>{t('specificationTitle')}</h2>
          <div className="spec-inputs">
            <div className="spec-input-group">
              <label>{t('uslLabel')}</label>
              <input
                type="number"
                step="0.01"
                placeholder="例如: 100"
                onChange={(e) => setUsl(parseFloat(e.target.value))}
              />
            </div>
            <div className="spec-input-group">
              <label>{t('lslLabel')}</label>
              <input
                type="number"
                step="0.01"
                placeholder="例如: 90"
                onChange={(e) => setLsl(parseFloat(e.target.value))}
              />
            </div>
          </div>
        </div>

        <div className="card">
          <button className="btn-primary" onClick={runAnalysis}>
            {t('runAnalysisButton')}
          </button>
        </div>

        {stats && (
          <div className="card">
            <h2>{t('resultsTitle')}</h2>
            <div className="results-grid">
              <div className="result-item">
                <div className="result-label">{t('n')}</div>
                <div className="result-value">{stats.n}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('mean')}</div>
                <div className="result-value">{formatNumber(stats.mean)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('stdDev')}</div>
                <div className="result-value">{formatNumber(stats.sigma)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('min')}</div>
                <div className="result-value">{formatNumber(stats.min)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('max')}</div>
                <div className="result-value">{formatNumber(stats.max)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('cp')}</div>
                <div className="result-value">{formatNumber(stats.cp)}</div>
              </div>
              <div
                className={`result-item ${
                  stats.cpk >= 1.67
                    ? 'excellent'
                    : stats.cpk >= 1.33
                    ? 'good'
                    : stats.cpk >= 1.0
                    ? 'marginal'
                    : 'poor'
                }`}
              >
                <div className="result-label">{t('cpk')}</div>
                <div className="result-value">{formatNumber(stats.cpk)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('pp')}</div>
                <div className="result-value">{formatNumber(stats.pp)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('ppk')}</div>
                <div className="result-value">{formatNumber(stats.ppk)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('sigmaLevel')}</div>
                <div className="result-value">{formatNumber(stats.sigmaLevel)}</div>
              </div>
              <div className="result-item">
                <div className="result-label">{t('ppm')}</div>
                <div className="result-value">{formatNumber(stats.ppm)}</div>
              </div>
            </div>
            <button className="btn-export" onClick={exportResults}>
              {t('exportResultsButton')}
            </button>
          </div>
        )}

        {stats && (
          <div>
            <h2 style={{ fontSize: '20px', fontWeight: '600', marginBottom: '24px', color: '#1f2937' }}>
              圖表分析
            </h2>
            <div className="charts-grid">
              <div className="chart-card">
                <h3>常態分佈曲線</h3>
                <div className="chart-container">
                  <canvas id="normalDistChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>直方圖</h3>
                <div className="chart-container">
                  <canvas id="histogramChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>X-bar 管制圖</h3>
                <div className="chart-container">
                  <canvas id="xbarChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>盒形圖</h3>
                <div className="chart-container">
                  <canvas id="boxPlotChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>趨勢圖</h3>
                <div className="chart-container">
                  <canvas id="trendChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>CPK 儀表板</h3>
                <div className="chart-container">
                  <canvas id="cpkGaugeChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>良率 & PPM</h3>
                <div className="chart-container">
                  <canvas id="yieldChart"></canvas>
                </div>
              </div>
              <div className="chart-card">
                <h3>Q-Q 圖</h3>
                <div className="chart-container">
                  <canvas id="qqPlotChart"></canvas>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </>
  );
}


