'use client';

import { Fragment, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  FileSpreadsheet,
  Layers3,
  Printer,
  UploadCloud,
  X,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';

type Row = Record<string, unknown>;
type Result = {
  objectClass: string;
  comObl: string;
  sag: string;
  functionalArea: string;
  dirDas: string;
  status: 'Matched' | 'Multiple matches' | 'Unmatched';
  months: number[];
  total: number;
};
type AgencyRow = {
  objectClass: string;
  comObl: string;
  sag: string;
  dirDas: string;
  months: number[];
  total: number;
};
type ProjectedExecutionRow = {
  requirement: string;
  mdep: string;
  sag: string;
  functionalArea: string;
  objectClass: string;
  comObl: string;
  dirDas: string;
  months: number[];
  total: number;
};
type DashboardRow = {
  dirDas: string;
  objectClass: string;
  comObl: string;
  months: number[];
};
type RollupRow = {
  dirDas: string;
  sag: string;
  months: number[];
  total: number;
};
type DashboardView =
  | 'dashboard'
  | 'projected'
  | 'sag'
  | 'consolidated'
  | 'schedule'
  | 'rollup';
const months = [
  'OCT',
  'NOV',
  'DEC',
  'JAN',
  'FEB',
  'MAR',
  'APR',
  'MAY',
  'JUN',
  'JUL',
  'AUG',
  'SEP',
];
const spendRequired = [
  'OBJECT CLASS',
  'COM / OBL',
  'SAG',
  'FUNCTIONAL AREA',
  'MDEP',
  ...months,
];
const apeRequired = ['Ape', 'MDEP', 'DIR/DAS'];

function text(value: unknown) {
  return String(value ?? '').trim();
}
function keyText(value: unknown) {
  return text(value).replace(/\s+/g, '').toUpperCase();
}
function compareText(a: string, b: string, direction: 'asc' | 'desc') {
  const order = a.localeCompare(b, undefined, { numeric: true });
  return direction === 'asc' ? order : -order;
}
function amount(value: unknown) {
  if (typeof value === 'number') return value;
  const raw = text(value).replace(/[$,\s]/g, '');
  if (!raw) return 0;
  const negative = raw.startsWith('(') && raw.endsWith(')');
  const parsed = Number(raw.replace(/[()]/g, ''));
  return Number.isFinite(parsed) ? (negative ? -parsed : parsed) : 0;
}
function currency(value: number) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  }).format(value);
}
function percent(value: number) {
  return new Intl.NumberFormat('en-US', {
    style: 'percent',
    maximumFractionDigits: 1,
  }).format(value);
}
function readRows(file: File, expectedColumns: string[]) {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    let bestRows: Row[] = [];
    let bestScore = -1;

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return;

      // Some Army-managed Excel builds leave the worksheet's !ref range at its
      // old size after rows are appended. Rebuild the range from the cells that
      // are actually present so newly added records are not silently ignored.
      const cellAddresses = Object.keys(sheet).filter(
        (address) => !address.startsWith('!'),
      );
      if (cellAddresses.length) {
        const decoded = cellAddresses.map((address) => XLSX.utils.decode_cell(address));
        const range = decoded.reduce(
          (current, cell) => ({
            s: {
              r: Math.min(current.s.r, cell.r),
              c: Math.min(current.s.c, cell.c),
            },
            e: {
              r: Math.max(current.e.r, cell.r),
              c: Math.max(current.e.c, cell.c),
            },
          }),
          { s: { r: Infinity, c: Infinity }, e: { r: 0, c: 0 } },
        );
        sheet['!ref'] = XLSX.utils.encode_range(range);
      }

      const rows = XLSX.utils.sheet_to_json<Row>(sheet, {
        defval: '',
        blankrows: false,
      });
      const present = new Set(Object.keys(rows[0] ?? {}).map(keyText));
      const score = expectedColumns.filter((column) =>
        present.has(keyText(column)),
      ).length;
      if (score > bestScore || (score === bestScore && rows.length > bestRows.length)) {
        bestRows = rows;
        bestScore = score;
      }
    });

    return bestRows;
  });
}
function missingColumns(rows: Row[], expected: string[]) {
  const present = new Set(Object.keys(rows[0] ?? {}).map(keyText));
  return expected.filter((column) => !present.has(keyText(column)));
}

export default function Home() {
  const spendInput = useRef<HTMLInputElement>(null);
  const apeInput = useRef<HTMLInputElement>(null);
  const [spendRows, setSpendRows] = useState<Row[]>([]);
  const [apeRows, setApeRows] = useState<Row[]>([]);
  const [spendName, setSpendName] = useState('');
  const [apeName, setApeName] = useState('');
  const [error, setError] = useState('');
  const [selectedComObl, setSelectedComObl] = useState<'COM' | 'OBL'>('COM');
  const [selectedSagComObl, setSelectedSagComObl] = useState<'COM' | 'OBL'>(
    'COM',
  );
  const [selectedDirDasComObl, setSelectedDirDasComObl] = useState<
    'COM' | 'OBL'
  >('COM');
  const [selectedSag, setSelectedSag] = useState('ALL');
  const [selectedDirDas, setSelectedDirDas] = useState('ALL');
  const [dashboardSortField, setDashboardSortField] = useState<
    'dirDas' | 'objectClass' | 'comObl'
  >('dirDas');
  const [dashboardSortDirection, setDashboardSortDirection] = useState<
    'asc' | 'desc'
  >('asc');
  const [projectedSortField, setProjectedSortField] = useState<
    'dirDas' | 'mdep' | 'sag' | 'functionalArea' | 'objectClass'
  >('dirDas');
  const [projectedSortDirection, setProjectedSortDirection] = useState<
    'asc' | 'desc'
  >('asc');
  const [sagSortField, setSagSortField] = useState<
    'dirDas' | 'objectClass' | 'comObl' | 'sag'
  >('dirDas');
  const [sagSortDirection, setSagSortDirection] = useState<'asc' | 'desc'>(
    'asc',
  );
  const [consolidatedSortField, setConsolidatedSortField] = useState<
    'dirDas' | 'objectClass' | 'comObl' | 'functionalArea' | 'status'
  >('dirDas');
  const [consolidatedSortDirection, setConsolidatedSortDirection] = useState<
    'asc' | 'desc'
  >('asc');
  const [scheduleGroupBy, setScheduleGroupBy] = useState<'sag' | 'dirDas'>(
    'dirDas',
  );
  const [scheduleDashboardGroupBy, setScheduleDashboardGroupBy] = useState<
    'sag' | 'dirDas' | 'objectClass' | 'functionalArea'
  >('sag');
  const [scheduleFundingType, setScheduleFundingType] = useState<'COM' | 'OBL'>(
    'COM',
  );
  const [schedulePeriodType, setSchedulePeriodType] = useState<
    'month' | 'quarter' | 'semiAnnual' | 'annual' | 'range'
  >('month');
  const [scheduleMonth, setScheduleMonth] = useState(0);
  const [scheduleQuarter, setScheduleQuarter] = useState(0);
  const [scheduleHalf, setScheduleHalf] = useState(0);
  const [scheduleStartMonth, setScheduleStartMonth] = useState(0);
  const [scheduleMonthCount, setScheduleMonthCount] = useState(3);
  const [rollupFundingType, setRollupFundingType] = useState<'COM' | 'OBL'>(
    'OBL',
  );
  const [rollupSortField, setRollupSortField] = useState<'dirDas' | 'sag'>(
    'dirDas',
  );
  const [rollupSortDirection, setRollupSortDirection] = useState<
    'asc' | 'desc'
  >('asc');
  const [dashboardDirDas, setDashboardDirDas] = useState('ALL');
  const [dashboardObjectClass, setDashboardObjectClass] = useState('ALL');
  const [dashboardFundingType, setDashboardFundingType] = useState<
    'ALL' | 'COM' | 'OBL'
  >('ALL');
  const [activeView, setActiveView] = useState<DashboardView>('dashboard');
  const [showSplash, setShowSplash] = useState(true);

  useEffect(() => {
    const timer = window.setTimeout(() => setShowSplash(false), 7000);
    return () => window.clearTimeout(timer);
  }, []);

  async function loadSpend(file?: File) {
    if (!file) return;
    try {
      const rows = await readRows(file, spendRequired);
      const missing = missingColumns(rows, spendRequired);
      if (missing.length)
        throw new Error(`Spend plan is missing: ${missing.join(', ')}`);
      setSpendRows(rows);
      setSpendName(file.name);
      setError('');
    } catch (cause) {
      setSpendRows([]);
      setSpendName('');
      setError(
        cause instanceof Error
          ? cause.message
          : 'The spend plan could not be read.',
      );
    }
  }
  async function loadApe(file?: File) {
    if (!file) return;
    try {
      const rows = await readRows(file, apeRequired);
      const missing = missingColumns(rows, apeRequired);
      if (missing.length)
        throw new Error(`APE reference is missing: ${missing.join(', ')}`);
      setApeRows(rows);
      setApeName(file.name);
      setError('');
    } catch (cause) {
      setApeRows([]);
      setApeName('');
      setError(
        cause instanceof Error
          ? cause.message
          : 'The APE reference could not be read.',
      );
    }
  }

  async function printCard(view: DashboardView) {
    const previousTitle = document.title;
    document
      .querySelectorAll('#print-root')
      .forEach((existingPrintRoot) => existingPrintRoot.remove());
    const source = document.querySelector<HTMLElement>(
      `[data-print-surface="${view}"]`,
    );
    if (!source) return;

    const printRoot = document.createElement('div');
    printRoot.id = 'print-root';
    printRoot.appendChild(source.cloneNode(true));
    document.body.appendChild(printRoot);

    const cleanup = () => {
      printRoot.remove();
      delete document.body.dataset.printView;
      document.title = previousTitle;
      window.removeEventListener('afterprint', cleanup);
    };
    document.body.dataset.printView = view;
    document.title = '';
    window.addEventListener('afterprint', cleanup);

    await document.fonts?.ready;
    await Promise.all(
      Array.from(printRoot.querySelectorAll('img')).map(async (image) => {
        if (image.complete) return;
        try {
          await image.decode();
        } catch {
          // Printing can continue with the browser's normal image fallback.
        }
      }),
    );
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
    await new Promise((resolve) => window.setTimeout(resolve, 250));
    window.print();
  }

  const fiscalYear = useMemo(() => {
    for (const row of spendRows) {
      const header = Object.keys(row).find((name) => keyText(name) === 'FUND');
      const match = text(header ? row[header] : '').match(/(\d{2})\D*$/);
      if (match) return `FY${match[1]}`;
    }
    return 'FY';
  }, [spendRows]);

  const results = useMemo<Result[]>(() => {
    if (!spendRows.length || !apeRows.length) return [];
    const field = (row: Row, name: string) =>
      row[
        Object.keys(row).find((header) => keyText(header) === keyText(name)) ??
          ''
      ];
    const map = new Map<string, Result>();
    spendRows.forEach((row) => {
      const objectClass = text(field(row, 'OBJECT CLASS')) || 'Unspecified';
      const comObl = text(field(row, 'COM / OBL')) || 'Unspecified';
      const sag = text(field(row, 'SAG')) || 'Unspecified';
      const functionalArea = keyText(field(row, 'FUNCTIONAL AREA'));
      const mdep = keyText(field(row, 'MDEP'));
      const candidates = apeRows.filter((ape) => {
        const apeCode = keyText(field(ape, 'Ape'));
        return (
          apeCode.length >= 6 &&
          functionalArea.startsWith(apeCode.slice(0, 6)) &&
          mdep === keyText(field(ape, 'MDEP'))
        );
      });
      const dirs = [
        ...new Set(
          candidates.map((ape) => text(field(ape, 'DIR/DAS'))).filter(Boolean),
        ),
      ];
      const status: Result['status'] =
        dirs.length === 1
          ? 'Matched'
          : dirs.length > 1
            ? 'Multiple matches'
            : 'Unmatched';
      const dirDas = dirs.length ? dirs.join(' / ') : 'No DIR/DASA match';
      const groupKey = [objectClass, comObl, sag, functionalArea, dirDas].join('|');
      const monthly = months.map((month) => amount(field(row, month)));
      const existing = map.get(groupKey) ?? {
        objectClass,
        comObl,
        sag,
        functionalArea,
        dirDas,
        status,
        months: Array(12).fill(0),
        total: 0,
      };
      existing.months = existing.months.map(
        (value, index) => value + monthly[index],
      );
      // The workbook's monthly amounts are running balances. September is the FY ending balance.
      existing.total = existing.months[existing.months.length - 1];
      map.set(groupKey, existing);
    });
    return [...map.values()].sort((a, b) => b.total - a.total);
  }, [spendRows, apeRows]);

  const filteredResults = useMemo(
    () => results.filter((row) => keyText(row.comObl) === selectedComObl),
    [results, selectedComObl],
  );
  const sortedFilteredResults = useMemo(
    () =>
      [...filteredResults].sort((a, b) => {
        const primary = compareText(
          a[consolidatedSortField],
          b[consolidatedSortField],
          consolidatedSortDirection,
        );
        return primary || compareText(a.dirDas, b.dirDas, 'asc');
      }),
    [filteredResults, consolidatedSortField, consolidatedSortDirection],
  );
  const sagOptions = useMemo(
    () =>
      [...new Set(results.map((row) => row.sag).filter(Boolean))].sort((a, b) =>
        a.localeCompare(b, undefined, { numeric: true }),
      ),
    [results],
  );
  const dirDasOptions = useMemo(
    () => [...new Set(results.map((row) => row.dirDas).filter(Boolean))].sort(),
    [results],
  );
  const sagResults = useMemo(
    () =>
      results.filter(
        (row) =>
          keyText(row.comObl) === selectedSagComObl &&
          (selectedSag === 'ALL' || row.sag === selectedSag),
      ),
    [results, selectedSagComObl, selectedSag],
  );
  const agencyRows = useMemo<AgencyRow[]>(() => {
    const map = new Map<string, AgencyRow>();
    sagResults.forEach((row) => {
      const groupKey = [row.objectClass, row.comObl, row.sag, row.dirDas].join(
        '|',
      );
      const existing = map.get(groupKey) ?? {
        objectClass: row.objectClass,
        comObl: row.comObl,
        sag: row.sag,
        dirDas: row.dirDas,
        months: Array(12).fill(0),
        total: 0,
      };
      existing.months = existing.months.map(
        (value, index) => value + row.months[index],
      );
      existing.total = existing.months[existing.months.length - 1];
      map.set(groupKey, existing);
    });
    return [...map.values()].sort((a, b) => {
      const primary = compareText(
        a[sagSortField],
        b[sagSortField],
        sagSortDirection,
      );
      return primary || compareText(a.dirDas, b.dirDas, 'asc');
    });
  }, [sagResults, sagSortField, sagSortDirection]);
  const projectedDirRows = useMemo<ProjectedExecutionRow[]>(() => {
    if (!spendRows.length || !apeRows.length) return [];
    const field = (row: Row, name: string) =>
      row[
        Object.keys(row).find((header) => keyText(header) === keyText(name)) ??
          ''
      ];
    return spendRows
      .flatMap((row) => {
        const comObl = text(field(row, 'COM / OBL')) || 'Unspecified';
        if (keyText(comObl) !== selectedDirDasComObl) return [];
        const functionalArea = keyText(field(row, 'FUNCTIONAL AREA'));
        const mdep = text(field(row, 'MDEP')) || 'Unspecified';
        const candidates = apeRows.filter((ape) => {
          const apeCode = keyText(field(ape, 'Ape'));
          return (
            apeCode.length >= 6 &&
            functionalArea.startsWith(apeCode.slice(0, 6)) &&
            keyText(mdep) === keyText(field(ape, 'MDEP'))
          );
        });
        const dirs = [
          ...new Set(
            candidates
              .map((ape) => text(field(ape, 'DIR/DAS')))
              .filter(Boolean),
          ),
        ];
        if (selectedDirDas !== 'ALL' && !dirs.includes(selectedDirDas))
          return [];
        const monthsValues = months.map((month) => amount(field(row, month)));
        return [
          {
            requirement:
              text(field(row, 'REQUIREMENT (or Contract Name)')) ||
              'Unspecified requirement',
            mdep,
            sag: text(field(row, 'SAG')) || 'Unspecified',
            functionalArea,
            objectClass: text(field(row, 'OBJECT CLASS')) || 'Unspecified',
            comObl,
            dirDas: dirs.join(' / ') || 'No DIR/DASA match',
            months: monthsValues,
            total: monthsValues[monthsValues.length - 1] ?? 0,
          },
        ];
      })
      .sort((a, b) => {
        const primary = compareText(
          a[projectedSortField],
          b[projectedSortField],
          projectedSortDirection,
        );
        return primary || compareText(a.dirDas, b.dirDas, 'asc');
      });
  }, [
    spendRows,
    apeRows,
    selectedDirDas,
    selectedDirDasComObl,
    projectedSortField,
    projectedSortDirection,
  ]);
  const dashboardRows = useMemo<DashboardRow[]>(() => {
    const map = new Map<string, DashboardRow>();
    results.forEach((row) => {
      const groupKey = [row.dirDas, row.objectClass, row.comObl].join('|');
      const existing = map.get(groupKey) ?? {
        dirDas: row.dirDas,
        objectClass: row.objectClass,
        comObl: row.comObl,
        months: Array(12).fill(0),
      };
      existing.months = existing.months.map(
        (value, index) => value + row.months[index],
      );
      map.set(groupKey, existing);
    });
    return [...map.values()].sort(
      (a, b) =>
        a.dirDas.localeCompare(b.dirDas) ||
        a.objectClass.localeCompare(b.objectClass) ||
        a.comObl.localeCompare(b.comObl),
    );
  }, [results]);
  const dashboardDirDasOptions = useMemo(
    () => [...new Set(dashboardRows.map((row) => row.dirDas))].sort(),
    [dashboardRows],
  );
  const dashboardObjectClassOptions = useMemo(
    () =>
      [...new Set(dashboardRows.map((row) => row.objectClass))].sort((a, b) =>
        a.localeCompare(b, undefined, { numeric: true }),
      ),
    [dashboardRows],
  );
  const visibleDashboardRows = useMemo(
    () =>
      dashboardRows
        .filter(
          (row) =>
            (dashboardDirDas === 'ALL' || row.dirDas === dashboardDirDas) &&
            (dashboardObjectClass === 'ALL' ||
              row.objectClass === dashboardObjectClass) &&
            (dashboardFundingType === 'ALL' ||
              keyText(row.comObl) === dashboardFundingType),
        )
        .sort((a, b) => {
          const primary = compareText(
            a[dashboardSortField],
            b[dashboardSortField],
            dashboardSortDirection,
          );
          return primary || compareText(a.dirDas, b.dirDas, 'asc');
        }),
    [
      dashboardRows,
      dashboardDirDas,
      dashboardObjectClass,
      dashboardFundingType,
      dashboardSortField,
      dashboardSortDirection,
    ],
  );
  const dashboardTotalRows = useMemo(
    () =>
      visibleDashboardRows.filter((row) =>
        dashboardFundingType === 'ALL'
          ? keyText(row.comObl) === 'OBL'
          : keyText(row.comObl) === dashboardFundingType,
      ),
    [visibleDashboardRows, dashboardFundingType],
  );
  const dashboardTotals = useMemo(
    () =>
      months.map((_, index) =>
        dashboardTotalRows.reduce((sum, row) => sum + row.months[index], 0),
      ),
    [dashboardTotalRows],
  );
  const dashboardFyTotal = dashboardTotals[dashboardTotals.length - 1] ?? 0;
  const scheduleMonthIndexes = useMemo(() => {
    if (schedulePeriodType === 'month') return [scheduleMonth];
    if (schedulePeriodType === 'quarter') {
      const start = scheduleQuarter * 3;
      return [start, start + 1, start + 2];
    }
    if (schedulePeriodType === 'semiAnnual') {
      const start = scheduleHalf * 6;
      return Array.from({ length: 6 }, (_, index) => start + index);
    }
    if (schedulePeriodType === 'annual') {
      return Array.from({ length: months.length }, (_, index) => index);
    }
    const count = Math.max(
      1,
      Math.min(scheduleMonthCount, months.length - scheduleStartMonth),
    );
    return Array.from({ length: count }, (_, index) => scheduleStartMonth + index);
  }, [
    schedulePeriodType,
    scheduleMonth,
    scheduleQuarter,
    scheduleHalf,
    scheduleStartMonth,
    scheduleMonthCount,
  ]);
  const scheduleRows = useMemo(() => {
    const grouped = new Map<string, number[]>();
    results
      .filter((row) => keyText(row.comObl) === scheduleFundingType)
      .forEach((row) => {
        const group = row[scheduleGroupBy] || 'Unspecified';
        const current = grouped.get(group) ?? Array(scheduleMonthIndexes.length).fill(0);
        scheduleMonthIndexes.forEach((monthIndex, displayIndex) => {
          const prior = monthIndex === 0 ? 0 : row.months[monthIndex - 1] ?? 0;
          current[displayIndex] += (row.months[monthIndex] ?? 0) - prior;
        });
        grouped.set(group, current);
      });
    return [...grouped.entries()]
      .map(([group, values]) => ({
        group,
        values,
        total: values.reduce((sum, value) => sum + value, 0),
      }))
      .sort((a, b) => compareText(a.group, b.group, 'asc'));
  }, [
    results,
    scheduleFundingType,
    scheduleGroupBy,
    scheduleMonthIndexes,
  ]);
  const scheduleTotals = useMemo(
    () =>
      scheduleMonthIndexes.map((_, index) =>
        scheduleRows.reduce((sum, row) => sum + row.values[index], 0),
      ),
    [scheduleRows, scheduleMonthIndexes],
  );
  const scheduleGrandTotal = scheduleTotals.reduce(
    (sum, value) => sum + value,
    0,
  );
  const scheduleDashboardRows = useMemo(() => {
    const grouped = new Map<string, { projected: number; annualRequirement: number }>();
    results
      .filter((row) => keyText(row.comObl) === scheduleFundingType)
      .forEach((row) => {
        const group = row[scheduleDashboardGroupBy] || 'Unspecified';
        const current = grouped.get(group) ?? { projected: 0, annualRequirement: 0 };
        scheduleMonthIndexes.forEach((monthIndex) => {
          const prior = monthIndex === 0 ? 0 : row.months[monthIndex - 1] ?? 0;
          current.projected += (row.months[monthIndex] ?? 0) - prior;
        });
        current.annualRequirement += row.months[months.length - 1] ?? 0;
        grouped.set(group, current);
      });
    return [...grouped.entries()]
      .map(([group, values]) => ({
        group,
        ...values,
        percentOfAnnual:
          values.annualRequirement === 0 ? 0 : values.projected / values.annualRequirement,
      }))
      .sort((a, b) => compareText(a.group, b.group, 'asc'));
  }, [results, scheduleFundingType, scheduleDashboardGroupBy, scheduleMonthIndexes]);
  const scheduleAnnualRequirement = scheduleDashboardRows.reduce(
    (sum, row) => sum + row.annualRequirement,
    0,
  );
  const schedulePercentOfAnnual =
    scheduleAnnualRequirement === 0 ? 0 : scheduleGrandTotal / scheduleAnnualRequirement;
  const rollupRows = useMemo<RollupRow[]>(() => {
    const grouped = new Map<string, RollupRow>();
    results
      .filter((row) => keyText(row.comObl) === rollupFundingType)
      .forEach((row) => {
        const groupKey = `${row.dirDas}|${row.sag}`;
        const current = grouped.get(groupKey) ?? {
          dirDas: row.dirDas,
          sag: row.sag,
          months: Array(months.length).fill(0),
          total: 0,
        };
        current.months = current.months.map(
          (value, index) => value + (row.months[index] ?? 0),
        );
        current.total = current.months[current.months.length - 1] ?? 0;
        grouped.set(groupKey, current);
      });
    return [...grouped.values()].sort((a, b) => {
      const primary = compareText(
        a[rollupSortField],
        b[rollupSortField],
        rollupSortDirection,
      );
      return primary || compareText(a.dirDas, b.dirDas, rollupSortDirection);
    });
  }, [results, rollupFundingType, rollupSortField, rollupSortDirection]);
  const rollupSections = useMemo(() => {
    const grouped = new Map<string, RollupRow[]>();
    rollupRows.forEach((row) => {
      const current = grouped.get(row.dirDas) ?? [];
      current.push(row);
      grouped.set(row.dirDas, current);
    });
    return [...grouped.entries()]
      .map(([dirDas, rows]) => ({
        dirDas,
        rows: [...rows].sort((a, b) =>
          compareText(a.sag, b.sag, rollupSortDirection),
        ),
        months: months.map((_, index) =>
          rows.reduce((sum, row) => sum + row.months[index], 0),
        ),
        total: rows.reduce((sum, row) => sum + row.total, 0),
      }))
      .sort((a, b) => compareText(a.dirDas, b.dirDas, rollupSortDirection));
  }, [rollupRows, rollupSortDirection]);
  const rollupAgencyBySag = useMemo(() => {
    const grouped = new Map<string, RollupRow>();
    rollupRows.forEach((row) => {
      const current = grouped.get(row.sag) ?? {
        dirDas: 'Agency',
        sag: row.sag,
        months: Array(months.length).fill(0),
        total: 0,
      };
      current.months = current.months.map(
        (value, index) => value + row.months[index],
      );
      current.total = current.months[current.months.length - 1] ?? 0;
      grouped.set(row.sag, current);
    });
    return [...grouped.values()].sort((a, b) =>
      compareText(a.sag, b.sag, rollupSortDirection),
    );
  }, [rollupRows, rollupSortDirection]);
  const rollupAgencyMonths = useMemo(
    () =>
      months.map((_, index) =>
        rollupRows.reduce((sum, row) => sum + row.months[index], 0),
      ),
    [rollupRows],
  );
  const rollupAgencyTotal =
    rollupAgencyMonths[rollupAgencyMonths.length - 1] ?? 0;
  const projectedTotals = useMemo(
    () =>
      months.map((_, index) =>
        projectedDirRows.reduce((sum, row) => sum + row.months[index], 0),
      ),
    [projectedDirRows],
  );
  const projectedFyTotal = projectedTotals[projectedTotals.length - 1] ?? 0;
  const totals = useMemo(
    () =>
      months.map((_, index) =>
        filteredResults.reduce((sum, row) => sum + row.months[index], 0),
      ),
    [filteredResults],
  );
  const sagTotals = useMemo(
    () =>
      months.map((_, index) =>
        sagResults.reduce((sum, row) => sum + row.months[index], 0),
      ),
    [sagResults],
  );
  const fyTotal = totals[totals.length - 1] ?? 0;
  const sagFyTotal = sagTotals[sagTotals.length - 1] ?? 0;
  const flagged = filteredResults.filter(
    (row) => row.status !== 'Matched',
  ).length;
  const ready = spendRows.length > 0 && apeRows.length > 0;
  const milestoneTarget = (index: number) =>
    index === 5 ? 0.5 : index === 8 ? 0.8 : null;
  const percentageTone = (values: number[], total: number, index: number) => {
    const target = milestoneTarget(index);
    if (target === null) return 'text-white/80';
    return total > 0 && (values[index] ?? 0) / total >= target
      ? 'bg-emerald-500/25 text-emerald-50'
      : 'bg-red-500/30 text-red-50';
  };
  const rollupPercentageTone = (
    values: number[],
    total: number,
    index: number,
  ) => {
    const target = milestoneTarget(index);
    if (target === null) return 'text-muted-foreground';
    return total > 0 && (values[index] ?? 0) / total >= target
      ? 'bg-emerald-100 text-emerald-900'
      : 'bg-red-100 text-red-900';
  };
  const milestoneShortfall = (
    value: number,
    total: number,
    index: number,
  ) => {
    const target = milestoneTarget(index);
    if (target === null || total <= 0) return null;
    const difference = value - total * target;
    return difference < 0 ? difference : null;
  };

  return (
    <main className="min-h-screen">
      {showSplash && (
        <div className="splash-screen">
          <div className="splash-content">
            <img
              className="splash-logo"
              src="/esd-logo.png"
              alt="Executive Services Division G-8 seal"
            />
            <div className="splash-title">
              Welcome to HQDA-G8 / ESD&apos;s
              <br />
              Agency SpendPlan Tool
            </div>
          </div>
        </div>
      )}
      <header className="border-b bg-[#18251d] text-white">
        <div className="mx-auto flex max-w-[1600px] items-center gap-3 px-5 py-4 sm:px-8">
          <img
            className="size-12 rounded-full border-2 border-[#d5a84f] object-cover"
            src="/esd-logo.png"
            alt="Executive Services Division G-8 seal"
          />
          <div>
            <h1 className="font-semibold">Agency SpendPlan Tool</h1>
            <p className="text-xs text-white/70">FY Consolidation SpendPlan</p>
            <p className="mt-0.5 text-xs text-white/60">
              Tool created by Mr. Sammy Payne, HQDA G-8/ESD
            </p>
          </div>
        </div>
      </header>
      <section className="mx-auto max-w-[1600px] px-5 py-8 sm:px-8">
        <div className="mb-7 max-w-3xl">
          <p className="text-sm font-medium text-primary">
            Workbook-driven consolidation
          </p>
          <h2 className="mt-1 text-3xl font-semibold tracking-tight">
            Turn a spend plan into an agency view.
          </h2>
          <p className="mt-3 text-muted-foreground">
            Upload the spend plan and the APE lookup. The tool consolidates
            Object Class, COM/OBL, Functional Area, DIR/DASA, and every month in
            your browser.
          </p>
        </div>
        <div className="grid gap-5 lg:grid-cols-2">
          <UploadCard
            title="1. Upload Spend Plan"
            description="Spend plan workbook"
            name={spendName}
            ready={!!spendRows.length}
            onClick={() => spendInput.current?.click()}
          />
          <UploadCard
            title="2. Upload APE Lookup"
            description="APE-to-DIR/DASA reference workbook"
            name={apeName}
            ready={!!apeRows.length}
            onClick={() => apeInput.current?.click()}
          />
        </div>
        <input
          ref={spendInput}
          className="sr-only"
          type="file"
          accept=".xlsx,.xls"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            event.currentTarget.value = '';
            void loadSpend(file);
          }}
        />
        <input
          ref={apeInput}
          className="sr-only"
          type="file"
          accept=".xlsx,.xls"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            event.currentTarget.value = '';
            void loadApe(file);
          }}
        />
        {error && (
          <div
            role="alert"
            className="mt-5 flex gap-3 rounded-xl border border-destructive/25 bg-destructive/10 p-4 text-sm text-destructive"
          >
            <AlertTriangle className="mt-0.5 size-4 shrink-0" />
            {error}
          </div>
        )}
        {ready && (
          <>
            <div className="mt-7 grid gap-4 sm:grid-cols-3">
              <Stat
                label={`${selectedComObl} consolidated records`}
                value={String(filteredResults.length)}
                icon={Layers3}
              />
              <Stat
                label="DIR/DASA review items"
                value={String(flagged)}
                icon={AlertTriangle}
              />
              <Stat
                label={`${fiscalYear} ${selectedComObl} total`}
                value={currency(fyTotal)}
                icon={FileSpreadsheet}
              />
            </div>
            {flagged > 0 && (
              <div className="mt-5 flex gap-3 rounded-xl border border-amber-500/30 bg-amber-50 p-4 text-sm text-amber-900">
                <AlertTriangle className="mt-0.5 size-4 shrink-0" />
                <span>
                  <strong>
                    {flagged} record{flagged === 1 ? '' : 's'}
                  </strong>{' '}
                  need review because the APE reference has no match or more
                  than one DIR/DASA candidate. They remain visible in the results
                  instead of being assigned automatically.
                </span>
              </div>
            )}
            <nav
              aria-label="Dashboard views"
              className="mt-6 flex flex-wrap gap-2 rounded-xl border bg-card p-3 print:hidden"
            >
              {(
                [
                  ['dashboard', 'Budget dashboard'],
                  ['projected', 'Projected execution'],
                  ['sag', 'View SAG'],
                  ['consolidated', 'Consolidated plan'],
                  ['schedule', 'Funding schedule'],
                  ['rollup', 'G-8 budget roll-up'],
                ] as const
              ).map(([view, label]) => (
                <Button
                  key={view}
                  variant={activeView === view ? 'default' : 'outline'}
                  size="sm"
                  onClick={() => setActiveView(view)}
                >
                  {label}
                </Button>
              ))}
            </nav>
            <p className="mt-2 text-xs text-muted-foreground print:hidden">
              Printing note: In the print dialog, expand{' '}
              <strong>More settings</strong> and turn off{' '}
              <strong>Headers and footers</strong>.
            </p>
            {activeView === 'dashboard' && (
              <div className="mt-5 flex flex-wrap items-end gap-3 rounded-xl border bg-card p-4 print:hidden">
                <label className="grid gap-1 text-sm font-medium">
                  DIR/DASA
                  <select
                    value={dashboardDirDas}
                    onChange={(event) => setDashboardDirDas(event.target.value)}
                    className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                  >
                    <option value="ALL">All DIR/DASA</option>
                    {dashboardDirDasOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="grid gap-1 text-sm font-medium">
                  Object Class
                  <select
                    value={dashboardObjectClass}
                    onChange={(event) =>
                      setDashboardObjectClass(event.target.value)
                    }
                    className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                  >
                    <option value="ALL">All Object Classes</option>
                    {dashboardObjectClassOptions.map((option) => (
                      <option key={option} value={option}>
                        {option}
                      </option>
                    ))}
                  </select>
                </label>
                <fieldset className="grid gap-1">
                  <legend className="text-sm font-medium">Funding type</legend>
                  <div
                    className="inline-flex rounded-md border bg-[#eef6ff] p-1"
                    aria-label="Select both, COM, or OBL"
                  >
                    {(['ALL', 'COM', 'OBL'] as const).map((value) => (
                      <Button
                        key={value}
                        type="button"
                        size="sm"
                        variant={
                          dashboardFundingType === value ? 'default' : 'ghost'
                        }
                        onClick={() => setDashboardFundingType(value)}
                      >
                        {value === 'ALL' ? 'Both' : value}
                      </Button>
                    ))}
                  </div>
                </fieldset>
                <label className="grid gap-1 text-sm font-medium">
                  Sort by
                  <select
                    value={dashboardSortField}
                    onChange={(event) =>
                      setDashboardSortField(
                        event.target.value as
                          | 'dirDas'
                          | 'objectClass'
                          | 'comObl',
                      )
                    }
                    className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                  >
                    <option value="dirDas">DIR/DASA</option>
                    <option value="objectClass">Object Class</option>
                    <option value="comObl">COM/OBL</option>
                  </select>
                </label>
                <Button
                  variant="outline"
                  onClick={() =>
                    setDashboardSortDirection((current) =>
                      current === 'asc' ? 'desc' : 'asc',
                    )
                  }
                >
                  {dashboardSortDirection === 'asc' ? 'A–Z' : 'Z–A'}
                </Button>
                <Button
                  className="ml-auto"
                  variant="outline"
                  onClick={() => printCard('dashboard')}
                >
                  <Printer /> Print dashboard
                </Button>
              </div>
            )}
            {activeView === 'dashboard' && (
              <section
                data-print-surface="dashboard"
                className="mt-6 overflow-hidden rounded-2xl border bg-card shadow-sm"
              >
                <div className="print-report-header border-b p-5">
                  <img
                    className="print-report-logo"
                    src="/esd-logo.png"
                    alt="Executive Services Directorate logo"
                  />
                  <h3 className="font-semibold">
                    {fiscalYear} Budget Execution Dashboard
                  </h3>
                  <p className="text-sm text-muted-foreground">
                    Agency-wide cumulative execution by DIR/DASA, Object Class,
                    COM/OBL, and month.
                  </p>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-[10px] xl:text-[11px]">
                    <thead className="bg-primary text-left text-primary-foreground">
                      <tr>
                        {['DIR/DASA', 'Object Class', 'COM/OBL', ...months].map(
                          (header) => (
                            <th
                              key={header}
                              className={`whitespace-nowrap border-b border-white/10 px-1.5 py-2.5 font-semibold ${months.includes(header) ? 'text-center' : 'text-left'}`}
                            >
                              {header}
                            </th>
                          ),
                        )}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleDashboardRows.map((row) => (
                        <tr
                          key={[row.dirDas, row.objectClass, row.comObl].join(
                            '|',
                          )}
                          className="border-b last:border-0 hover:bg-muted/30"
                        >
                          <td className="whitespace-nowrap px-1.5 py-2.5 font-medium">
                            {row.dirDas}
                          </td>
                          <td className="whitespace-nowrap px-1.5 py-2.5 font-medium">
                            {row.objectClass}
                          </td>
                          <td className="whitespace-nowrap px-1.5 py-2.5">
                            {row.comObl}
                          </td>
                          {row.months.map((value, index) => (
                            <td
                              key={index}
                              className="whitespace-nowrap px-1.5 py-2.5 text-right tabular-nums"
                            >
                              {currency(value)}
                            </td>
                          ))}
                        </tr>
                      ))}
                      <tr className="bg-[#142541] text-white">
                        <td colSpan={3} className="px-1.5 py-2.5 font-semibold">
                          {dashboardDirDas === 'ALL'
                            ? 'Agency total'
                            : `${dashboardDirDas} total`}
                        </td>
                        {dashboardTotals.map((value, index) => (
                          <td
                            key={index}
                            className="whitespace-nowrap px-1.5 py-2.5 text-right font-semibold tabular-nums"
                          >
                            {currency(value)}
                          </td>
                        ))}
                      </tr>
                      <tr className="bg-[#20395e] text-white">
                        <td colSpan={3} className="px-1.5 py-2.5 font-semibold">
                          Percent of{' '}
                          {dashboardFundingType === 'ALL'
                            ? 'OBL'
                            : dashboardFundingType}{' '}
                          total
                        </td>
                        {dashboardTotals.map((value, index) => (
                          <td
                            key={index}
                            className={`whitespace-nowrap px-1.5 py-2.5 text-right font-semibold tabular-nums ${percentageTone(dashboardTotals, dashboardFyTotal, index)}`}
                          >
                            {percent(
                              dashboardFyTotal ? value / dashboardFyTotal : 0,
                            )}
                            {milestoneTarget(index) !== null && (
                              <span className="ml-1 text-[10px] opacity-75">
                                /{percent(milestoneTarget(index) ?? 0)}
                              </span>
                            )}
                            {milestoneShortfall(
                              value,
                              dashboardFyTotal,
                              index,
                            ) !== null && (
                              <span className="mt-0.5 block text-[10px]">
                                {currency(
                                  milestoneShortfall(
                                    value,
                                    dashboardFyTotal,
                                    index,
                                  ) ?? 0,
                                )}
                              </span>
                            )}
                          </td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>
              </section>
            )}
            {activeView === 'projected' && (
              <div className="mt-5 flex flex-wrap items-end justify-between gap-3 rounded-xl border bg-card p-4 print:hidden">
                <div className="flex flex-wrap items-end gap-3">
                  <label className="grid gap-1 text-sm font-medium">
                    DIR/DASA
                    <select
                      value={selectedDirDas}
                      onChange={(event) =>
                        setSelectedDirDas(event.target.value)
                      }
                      className="h-10 min-w-36 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="ALL">All DIR/DASA</option>
                      {dirDasOptions.map((option) => (
                        <option key={option} value={option}>
                          {option}
                        </option>
                      ))}
                    </select>
                  </label>
                  <FundingTypeButtons
                    value={selectedDirDasComObl}
                    onChange={setSelectedDirDasComObl}
                  />
                  <label className="grid gap-1 text-sm font-medium">
                    Sort by
                    <select
                      value={projectedSortField}
                      onChange={(event) =>
                        setProjectedSortField(
                          event.target.value as
                            | 'dirDas'
                            | 'mdep'
                            | 'sag'
                            | 'functionalArea'
                            | 'objectClass',
                        )
                      }
                      className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="dirDas">DIR/DASA</option>
                      <option value="mdep">MDEP</option>
                      <option value="sag">SAG</option>
                      <option value="functionalArea">Functional Area</option>
                      <option value="objectClass">Object Class</option>
                    </select>
                  </label>
                  <Button
                    variant="outline"
                    onClick={() =>
                      setProjectedSortDirection((current) =>
                        current === 'asc' ? 'desc' : 'asc',
                      )
                    }
                  >
                    {projectedSortDirection === 'asc' ? 'A–Z' : 'Z–A'}
                  </Button>
                </div>
                <Button
                  variant="outline"
                  onClick={() => printCard('projected')}
                >
                  <Printer /> Print{' '}
                  {activeView === 'projected'
                    ? 'projected execution'
                    : 'DIR/DASA'}
                </Button>
              </div>
            )}
            {activeView === 'sag' && (
              <div className="mt-5 flex flex-wrap items-end justify-between gap-3 rounded-xl border bg-card p-4 print:hidden">
                <div className="flex flex-wrap items-end gap-3">
                  <label className="grid gap-1 text-sm font-medium">
                    SAG
                    <select
                      value={selectedSag}
                      onChange={(event) => setSelectedSag(event.target.value)}
                      className="h-10 min-w-28 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="ALL">All SAGs</option>
                      {sagOptions.map((option) => (
                        <option key={option} value={option}>
                          {option}
                        </option>
                      ))}
                    </select>
                  </label>
                  <FundingTypeButtons
                    value={selectedSagComObl}
                    onChange={setSelectedSagComObl}
                  />
                  <label className="grid gap-1 text-sm font-medium">
                    Sort by
                    <select
                      value={sagSortField}
                      onChange={(event) =>
                        setSagSortField(
                          event.target.value as
                            | 'dirDas'
                            | 'objectClass'
                            | 'comObl'
                            | 'sag',
                        )
                      }
                      className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="dirDas">DIR/DASA</option>
                      <option value="objectClass">Object Class</option>
                      <option value="comObl">COM/OBL</option>
                      <option value="sag">SAG</option>
                    </select>
                  </label>
                  <Button
                    variant="outline"
                    onClick={() =>
                      setSagSortDirection((current) =>
                        current === 'asc' ? 'desc' : 'asc',
                      )
                    }
                  >
                    {sagSortDirection === 'asc' ? 'A–Z' : 'Z–A'}
                  </Button>
                </div>
                <Button variant="outline" onClick={() => printCard('sag')}>
                  <Printer /> Print SAG
                </Button>
              </div>
            )}
            {activeView === 'consolidated' && (
              <div className="mt-5 flex flex-wrap items-end justify-between gap-3 rounded-xl border bg-card p-4 print:hidden">
                <div className="flex flex-wrap items-end gap-3">
                  <FundingTypeButtons
                    value={selectedComObl}
                    onChange={setSelectedComObl}
                  />
                  <label className="grid gap-1 text-sm font-medium">
                    Sort by
                    <select
                      value={consolidatedSortField}
                      onChange={(event) =>
                        setConsolidatedSortField(
                          event.target.value as
                            | 'dirDas'
                            | 'objectClass'
                            | 'comObl'
                            | 'functionalArea'
                            | 'status',
                        )
                      }
                      className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="dirDas">DIR/DASA</option>
                      <option value="objectClass">Object Class</option>
                      <option value="comObl">COM/OBL</option>
                      <option value="functionalArea">Functional Area</option>
                      <option value="status">Status</option>
                    </select>
                  </label>
                  <Button
                    variant="outline"
                    onClick={() =>
                      setConsolidatedSortDirection((current) =>
                        current === 'asc' ? 'desc' : 'asc',
                      )
                    }
                  >
                    {consolidatedSortDirection === 'asc' ? 'A–Z' : 'Z–A'}
                  </Button>
                </div>
                <Button
                  variant="outline"
                  onClick={() => printCard('consolidated')}
                >
                  <Printer /> Print consolidated
                </Button>
              </div>
            )}
            {activeView === 'schedule' && (
              <div className="mt-5 flex flex-wrap items-end justify-between gap-3 rounded-xl border bg-card p-4 print:hidden">
                <div className="flex flex-wrap items-end gap-3">
                  <label className="grid gap-1 text-sm font-medium">
                    Group by
                    <select
                      value={scheduleGroupBy}
                      onChange={(event) =>
                        setScheduleGroupBy(
                          event.target.value as 'sag' | 'dirDas',
                        )
                      }
                      className="h-10 min-w-36 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="dirDas">DIR/DASA</option>
                      <option value="sag">SAG</option>
                    </select>
                  </label>
                  <FundingTypeButtons
                    value={scheduleFundingType}
                    onChange={setScheduleFundingType}
                  />
                  <label className="grid gap-1 text-sm font-medium">
                    Period
                    <select
                      value={schedulePeriodType}
                      onChange={(event) =>
                        setSchedulePeriodType(
                          event.target.value as
                            | 'month'
                            | 'quarter'
                            | 'semiAnnual'
                            | 'annual'
                            | 'range',
                        )
                      }
                      className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="month">One month</option>
                      <option value="quarter">Quarter</option>
                      <option value="semiAnnual">Semi-Annual</option>
                      <option value="annual">Annual</option>
                      <option value="range">Number of months</option>
                    </select>
                  </label>
                  {schedulePeriodType === 'month' && (
                    <label className="grid gap-1 text-sm font-medium">
                      Month
                      <select
                        value={scheduleMonth}
                        onChange={(event) =>
                          setScheduleMonth(Number(event.target.value))
                        }
                        className="h-10 min-w-28 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                      >
                        {months.map((month, index) => (
                          <option key={month} value={index}>
                            {month}
                          </option>
                        ))}
                      </select>
                    </label>
                  )}
                  {schedulePeriodType === 'quarter' && (
                    <label className="grid gap-1 text-sm font-medium">
                      Quarter
                      <select
                        value={scheduleQuarter}
                        onChange={(event) =>
                          setScheduleQuarter(Number(event.target.value))
                        }
                        className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                      >
                        <option value={0}>Q1 (OCT–DEC)</option>
                        <option value={1}>Q2 (JAN–MAR)</option>
                        <option value={2}>Q3 (APR–JUN)</option>
                        <option value={3}>Q4 (JUL–SEP)</option>
                      </select>
                    </label>
                  )}
                  {schedulePeriodType === 'semiAnnual' && (
                    <label className="grid gap-1 text-sm font-medium">
                      Semi-Annual period
                      <select
                        value={scheduleHalf}
                        onChange={(event) =>
                          setScheduleHalf(Number(event.target.value))
                        }
                        className="h-10 min-w-44 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                      >
                        <option value={0}>First Half (OCT–MAR)</option>
                        <option value={1}>Second Half (APR–SEP)</option>
                      </select>
                    </label>
                  )}
                  {schedulePeriodType === 'range' && (
                    <>
                      <label className="grid gap-1 text-sm font-medium">
                        Starting month
                        <select
                          value={scheduleStartMonth}
                          onChange={(event) =>
                            setScheduleStartMonth(Number(event.target.value))
                          }
                          className="h-10 min-w-32 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                        >
                          {months.map((month, index) => (
                            <option key={month} value={index}>
                              {month}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label className="grid gap-1 text-sm font-medium">
                        Number of months
                        <select
                          value={Math.min(
                            scheduleMonthCount,
                            months.length - scheduleStartMonth,
                          )}
                          onChange={(event) =>
                            setScheduleMonthCount(Number(event.target.value))
                          }
                          className="h-10 min-w-28 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                        >
                          {Array.from(
                            { length: months.length - scheduleStartMonth },
                            (_, index) => index + 1,
                          ).map((count) => (
                            <option key={count} value={count}>
                              {count}
                            </option>
                          ))}
                        </select>
                      </label>
                    </>
                  )}
                </div>
                <Button variant="outline" onClick={() => printCard('schedule')}>
                  <Printer /> Print funding schedule
                </Button>
              </div>
            )}
            <section
              data-print-surface="schedule"
              className={`${activeView === 'schedule' ? 'block' : 'hidden'} mt-6 overflow-hidden rounded-2xl border bg-card shadow-sm`}
            >
              <div className="print-report-header border-b p-5">
                <img
                  className="print-report-logo"
                  src="/esd-logo.png"
                  alt="Executive Services Directorate logo"
                />
                <h3 className="font-semibold">
                  {fiscalYear} Planned Funding Schedule by{' '}
                  {scheduleGroupBy === 'dirDas' ? 'DIR/DASA' : 'SAG'}
                </h3>
                <p className="text-sm text-muted-foreground">
                  Planned {scheduleFundingType} receipts are calculated as the
                  change in the cumulative Spend Plan balance for each month.
                </p>
              </div>
              <div className="border-b bg-muted/20 p-5">
                <div className="mb-4 flex flex-wrap items-end justify-between gap-3 print:hidden">
                  <div>
                    <h4 className="font-semibold">Funding schedule dashboard</h4>
                    <p className="text-sm text-muted-foreground">
                      Selected-period projected receipts compared with the annual requirement.
                    </p>
                  </div>
                  <label className="grid gap-1 text-sm font-medium">
                    Dashboard breakdown
                    <select
                      value={scheduleDashboardGroupBy}
                      onChange={(event) =>
                        setScheduleDashboardGroupBy(
                          event.target.value as
                            | 'sag'
                            | 'dirDas'
                            | 'objectClass'
                            | 'functionalArea',
                        )
                      }
                      className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                    >
                      <option value="sag">SAG</option>
                      <option value="dirDas">DIR/DASA</option>
                      <option value="objectClass">Object Class</option>
                      <option value="functionalArea">Functional Area</option>
                    </select>
                  </label>
                </div>
                <div className="grid gap-3 sm:grid-cols-3">
                  <div className="rounded-xl border bg-card p-4">
                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Projected to receive</p>
                    <p className="mt-1 text-xl font-bold tabular-nums">{currency(scheduleGrandTotal)}</p>
                  </div>
                  <div className="rounded-xl border bg-card p-4">
                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Annual requirement</p>
                    <p className="mt-1 text-xl font-bold tabular-nums">{currency(scheduleAnnualRequirement)}</p>
                  </div>
                  <div className="rounded-xl border bg-card p-4">
                    <p className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">Percent of annual requirement</p>
                    <p className="mt-1 text-xl font-bold tabular-nums">{percent(schedulePercentOfAnnual)}</p>
                  </div>
                </div>
                <div className="mt-4 overflow-x-auto rounded-xl border bg-card">
                  <table className="w-full min-w-[620px] text-sm">
                    <thead className="bg-[#142541] text-white">
                      <tr>
                        <th className="px-3 py-2 text-left font-semibold">
                          {scheduleDashboardGroupBy === 'dirDas'
                            ? 'DIR/DASA'
                            : scheduleDashboardGroupBy === 'objectClass'
                              ? 'Object Class'
                              : scheduleDashboardGroupBy === 'functionalArea'
                                ? 'Functional Area'
                                : 'SAG'}
                        </th>
                        <th className="px-3 py-2 text-center font-semibold">Projected to receive</th>
                        <th className="px-3 py-2 text-center font-semibold">Annual requirement</th>
                        <th className="px-3 py-2 text-center font-semibold">Percent of annual</th>
                      </tr>
                    </thead>
                    <tbody>
                      {scheduleDashboardRows.map((row) => (
                        <tr key={row.group} className="border-b last:border-0">
                          <td className="px-3 py-2 font-medium">{row.group}</td>
                          <td className="whitespace-nowrap px-3 py-2 text-center tabular-nums">{currency(row.projected)}</td>
                          <td className="whitespace-nowrap px-3 py-2 text-center tabular-nums">{currency(row.annualRequirement)}</td>
                          <td className="whitespace-nowrap px-3 py-2 text-center font-semibold tabular-nums">{percent(row.percentOfAnnual)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[700px] text-sm">
                  <thead className="bg-primary text-primary-foreground">
                    <tr>
                      <th className="border-b border-white/10 px-3 py-3 text-left font-semibold">
                        {scheduleGroupBy === 'dirDas' ? 'DIR/DASA' : 'SAG'}
                      </th>
                      {scheduleMonthIndexes.map((monthIndex) => (
                        <th
                          key={monthIndex}
                          className="whitespace-nowrap border-b border-white/10 px-3 py-3 text-center font-semibold"
                        >
                          {months[monthIndex]}
                        </th>
                      ))}
                      <th className="whitespace-nowrap border-b border-white/10 px-3 py-3 text-center font-semibold">
                        Selected period total
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {scheduleRows.map((row) => (
                      <tr key={row.group} className="border-b last:border-0">
                        <td className="px-3 py-3 font-medium">{row.group}</td>
                        {row.values.map((value, index) => (
                          <td
                            key={scheduleMonthIndexes[index]}
                            className="whitespace-nowrap px-3 py-3 text-center tabular-nums"
                          >
                            {currency(value)}
                          </td>
                        ))}
                        <td className="whitespace-nowrap px-3 py-3 text-center font-semibold tabular-nums">
                          {currency(row.total)}
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-[#142541] text-white">
                      <td className="px-3 py-3 font-semibold">Agency total</td>
                      {scheduleTotals.map((value, index) => (
                        <td
                          key={scheduleMonthIndexes[index]}
                          className="whitespace-nowrap px-3 py-3 text-center font-semibold tabular-nums"
                        >
                          {currency(value)}
                        </td>
                      ))}
                      <td className="whitespace-nowrap px-3 py-3 text-center font-semibold tabular-nums">
                        {currency(scheduleGrandTotal)}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
            {activeView === 'rollup' && (
              <div className="mt-5 flex flex-wrap items-end gap-3 rounded-xl border bg-card p-4 print:hidden">
                <FundingTypeButtons
                  value={rollupFundingType}
                  onChange={setRollupFundingType}
                />
                <label className="grid gap-1 text-sm font-medium">
                  Sort by
                  <select
                    value={rollupSortField}
                    onChange={(event) =>
                      setRollupSortField(
                        event.target.value as 'dirDas' | 'sag',
                      )
                    }
                    className="h-10 min-w-40 rounded-md border bg-[#eef6ff] px-3 text-sm font-medium shadow-sm outline-none focus:ring-2 focus:ring-primary"
                  >
                    <option value="dirDas">DIR/DASA</option>
                    <option value="sag">SAG</option>
                  </select>
                </label>
                <Button
                  variant="outline"
                  onClick={() =>
                    setRollupSortDirection((current) =>
                      current === 'asc' ? 'desc' : 'asc',
                    )
                  }
                >
                  {rollupSortDirection === 'asc' ? 'A–Z' : 'Z–A'}
                </Button>
                <Button
                  className="ml-auto"
                  variant="outline"
                  onClick={() => printCard('rollup')}
                >
                  <Printer /> Print G-8 roll-up
                </Button>
              </div>
            )}
            <section
              data-print-surface="rollup"
              className={`${activeView === 'rollup' ? 'block' : 'hidden'} mt-6 overflow-hidden rounded-2xl border bg-card shadow-sm`}
            >
              <div className="print-report-header border-b p-5">
                <img
                  className="print-report-logo"
                  src="/esd-logo.png"
                  alt="Executive Services Directorate logo"
                />
                <h3 className="font-semibold">
                  {fiscalYear} G-8 Budget Roll-Up — {rollupFundingType}
                </h3>
                <p className="text-sm text-muted-foreground">
                  Cumulative monthly execution by DIR/DASA and SAG. Percentages
                  compare each month with the September annual requirement.
                </p>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[2200px] text-xs">
                  <thead className="bg-primary text-primary-foreground">
                    <tr>
                      <th
                        rowSpan={2}
                        className="border-b border-white/10 px-3 py-3 text-left font-semibold"
                      >
                        DIR/DASA
                      </th>
                      <th
                        rowSpan={2}
                        className="border-b border-white/10 px-3 py-3 text-center font-semibold"
                      >
                        SAG
                      </th>
                      {months.map((month) => (
                        <th
                          key={month}
                          colSpan={2}
                          className="border-b border-l border-white/10 px-2 py-2 text-center font-semibold"
                        >
                          {month}
                        </th>
                      ))}
                      <th
                        rowSpan={2}
                        className="border-b border-l border-white/10 px-3 py-3 text-center font-semibold"
                      >
                        Annual requirement
                      </th>
                    </tr>
                    <tr>
                      {months.flatMap((month) => [
                        <th
                          key={`${month}-amount`}
                          className="border-b border-l border-white/10 px-2 py-2 text-center font-semibold"
                        >
                          Amount
                        </th>,
                        <th
                          key={`${month}-percent`}
                          className="border-b border-white/10 px-2 py-2 text-center font-semibold"
                        >
                          %
                        </th>,
                      ])}
                    </tr>
                  </thead>
                  <tbody>
                    {rollupSections.map((section) => (
                      <Fragment key={section.dirDas}>
                        <tr className="bg-muted/40">
                          <td
                            colSpan={27}
                            className="px-3 py-2 font-semibold text-foreground"
                          >
                            {section.dirDas} Funding Roll-Up
                          </td>
                        </tr>
                        {section.rows.map((row) => (
                          <tr
                            key={`${row.dirDas}-${row.sag}`}
                            className="border-b last:border-0"
                          >
                            <td className="px-3 py-2 font-medium">{row.dirDas}</td>
                            <td className="px-3 py-2 text-center">{row.sag}</td>
                            {row.months.flatMap((value, index) => {
                              const rowPercent = row.total > 0 ? value / row.total : 0;
                              const shortfall = milestoneShortfall(value, row.total, index);
                              return [
                                <td
                                  key={`${index}-amount`}
                                  className="whitespace-nowrap border-l px-2 py-2 text-center tabular-nums"
                                >
                                  {currency(value)}
                                </td>,
                                <td
                                  key={`${index}-percent`}
                                  className={`whitespace-nowrap px-2 py-2 text-center tabular-nums ${rollupPercentageTone(row.months, row.total, index)}`}
                                >
                                  <span className="font-semibold">{percent(rowPercent)}</span>
                                  {shortfall !== null && (
                                    <span className="block whitespace-nowrap text-[9px] font-medium">
                                      {currency(shortfall)}
                                    </span>
                                  )}
                                </td>,
                              ];
                            })}
                            <td className="whitespace-nowrap border-l px-3 py-2 text-center font-semibold tabular-nums">
                              {currency(row.total)}
                            </td>
                          </tr>
                        ))}
                        <tr className="bg-[#e7f4ea] font-semibold">
                          <td className="px-3 py-2">{section.dirDas} total</td>
                          <td className="px-3 py-2 text-center">All SAGs</td>
                          {section.months.flatMap((value, index) => [
                            <td
                              key={`${index}-section-amount`}
                              className="whitespace-nowrap border-l px-2 py-2 text-center tabular-nums"
                            >
                              {currency(value)}
                            </td>,
                            <td
                              key={`${index}-section-percent`}
                              className={`whitespace-nowrap px-2 py-2 text-center tabular-nums ${rollupPercentageTone(section.months, section.total, index)}`}
                            >
                              {percent(section.total > 0 ? value / section.total : 0)}
                            </td>,
                          ])}
                          <td className="whitespace-nowrap border-l px-3 py-2 text-center tabular-nums">
                            {currency(section.total)}
                          </td>
                        </tr>
                      </Fragment>
                    ))}
                    <tr className="bg-[#142541] text-white">
                      <td
                        colSpan={27}
                        className="px-3 py-2 font-semibold"
                      >
                        G-8 Total(s) by SAG
                      </td>
                    </tr>
                    {rollupAgencyBySag.map((row) => (
                      <tr key={`agency-${row.sag}`} className="border-b">
                        <td className="px-3 py-2 font-medium">Agency</td>
                        <td className="px-3 py-2 text-center">{row.sag}</td>
                        {row.months.flatMap((value, index) => [
                          <td
                            key={`${index}-agency-sag-amount`}
                            className="whitespace-nowrap border-l px-2 py-2 text-center tabular-nums"
                          >
                            {currency(value)}
                          </td>,
                          <td
                            key={`${index}-agency-sag-percent`}
                            className={`whitespace-nowrap px-2 py-2 text-center tabular-nums ${rollupPercentageTone(row.months, row.total, index)}`}
                          >
                            {percent(row.total > 0 ? value / row.total : 0)}
                          </td>,
                        ])}
                        <td className="whitespace-nowrap border-l px-3 py-2 text-center font-semibold tabular-nums">
                          {currency(row.total)}
                        </td>
                      </tr>
                    ))}
                    <tr className="bg-[#142541] text-white">
                      <td className="px-3 py-3 font-semibold">Agency total</td>
                      <td className="px-3 py-3 text-center font-semibold">All SAGs</td>
                      {rollupAgencyMonths.flatMap((value, index) => [
                        <td
                          key={`${index}-grand-amount`}
                          className="whitespace-nowrap border-l border-white/10 px-2 py-3 text-center font-semibold tabular-nums"
                        >
                          {currency(value)}
                        </td>,
                        <td
                          key={`${index}-grand-percent`}
                          className={`whitespace-nowrap px-2 py-3 text-center font-semibold tabular-nums ${percentageTone(rollupAgencyMonths, rollupAgencyTotal, index)}`}
                        >
                          {percent(rollupAgencyTotal > 0 ? value / rollupAgencyTotal : 0)}
                        </td>,
                      ])}
                      <td className="whitespace-nowrap border-l border-white/10 px-3 py-3 text-center font-semibold tabular-nums">
                        {currency(rollupAgencyTotal)}
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
            <section
              data-print-surface="projected"
              className={`${activeView === 'projected' ? 'block' : 'hidden'} mt-6 overflow-hidden rounded-2xl border bg-card shadow-sm`}
            >
              <div className="print-report-header border-b p-5">
                <img
                  className="print-report-logo"
                  src="/esd-logo.png"
                  alt="Executive Services Directorate logo"
                />
                <h3 className="font-semibold">
                  {fiscalYear} Projected Execution by DIR/DASA
                </h3>
                <p className="text-sm text-muted-foreground">
                  Requirement-level projected execution for{' '}
                  {selectedDirDas === 'ALL'
                    ? 'all DIR/DASA organizations'
                    : selectedDirDas}
                  , filtered to {selectedDirDasComObl}.
                </p>
              </div>
              <div className="overflow-x-auto">
                <table className="min-w-[2450px] w-full text-sm">
                  <thead className="bg-muted/60 text-left">
                    <tr>
                      {[
                        'DIR/DASA',
                        'Requirement / Contract Name',
                        'MDEP',
                        'SAG',
                        'Functional Area',
                        'Object Class',
                        'COM/OBL',
                        ...months,
                      ].map((header) => (
                        <th
                          key={header}
                          className={`whitespace-nowrap border-b px-3 py-3 font-semibold ${months.includes(header) ? 'text-center' : 'text-left'}`}
                        >
                          {header}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {projectedDirRows.map((row, index) => (
                      <tr
                        key={[
                          row.requirement,
                          row.mdep,
                          row.functionalArea,
                          row.comObl,
                          index,
                        ].join('|')}
                        className="border-b last:border-0 hover:bg-muted/30"
                      >
                        <td className="px-3 py-3">{row.dirDas}</td>
                        <td className="max-w-80 px-3 py-3 font-medium whitespace-normal">
                          {row.requirement}
                        </td>
                        <td className="px-3 py-3 font-mono">{row.mdep}</td>
                        <td className="px-3 py-3 font-mono">{row.sag}</td>
                        <td className="px-3 py-3 font-mono">
                          {row.functionalArea}
                        </td>
                        <td className="px-3 py-3">{row.objectClass}</td>
                        <td className="px-3 py-3">{row.comObl}</td>
                        {row.months.map((value, monthIndex) => (
                          <td
                            key={monthIndex}
                            className="px-3 py-3 text-right tabular-nums"
                          >
                            {currency(value)}
                          </td>
                        ))}
                      </tr>
                    ))}
                    {projectedDirRows.length > 0 && (
                      <>
                        <tr className="bg-[#142541] text-white">
                          <td colSpan={7} className="px-3 py-3 font-semibold">
                            {selectedDirDas === 'ALL'
                              ? `${selectedDirDasComObl} agency total`
                              : `${selectedDirDas} ${selectedDirDasComObl} total`}
                          </td>
                          {projectedTotals.map((value, index) => (
                            <td
                              key={index}
                              className="px-3 py-3 text-right font-semibold tabular-nums"
                            >
                              {currency(value)}
                            </td>
                          ))}
                        </tr>
                        <tr className="bg-[#20395e] text-white">
                          <td colSpan={7} className="px-3 py-3 font-semibold">
                            Percent of {fiscalYear} {selectedDirDasComObl} total
                          </td>
                          {projectedTotals.map((value, index) => (
                            <td
                              key={index}
                              className={`px-3 py-3 text-right font-semibold tabular-nums ${percentageTone(projectedTotals, projectedFyTotal, index)}`}
                            >
                              {percent(
                                projectedFyTotal ? value / projectedFyTotal : 0,
                              )}
                              {milestoneTarget(index) !== null && (
                                <span className="ml-1 text-[10px] opacity-75">
                                  /{percent(milestoneTarget(index) ?? 0)}
                                </span>
                              )}
                              {milestoneShortfall(
                                value,
                                projectedFyTotal,
                                index,
                              ) !== null && (
                                <span className="milestone-shortfall mt-0.5 block whitespace-nowrap leading-none">
                                  {currency(
                                    milestoneShortfall(
                                      value,
                                      projectedFyTotal,
                                      index,
                                    ) ?? 0,
                                  )}
                                </span>
                              )}
                            </td>
                          ))}
                        </tr>
                      </>
                    )}
                    {projectedDirRows.length === 0 && (
                      <tr>
                        <td
                          colSpan={19}
                          className="px-3 py-8 text-center text-muted-foreground"
                        >
                          No projected execution records match the current
                          DIR/DASA and funding type selections.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </section>
            <section
              data-print-surface="sag"
              className={`${activeView === 'sag' ? 'block' : 'hidden'} mt-6 overflow-hidden rounded-2xl border bg-card shadow-sm`}
            >
              <div className="print-report-header border-b p-5">
                <img
                  className="print-report-logo"
                  src="/esd-logo.png"
                  alt="Executive Services Directorate logo"
                />
                <h3 className="font-semibold">
                  {fiscalYear} Agency View by SAG
                </h3>
                <p className="text-sm text-muted-foreground">
                  A high-level {selectedSagComObl} view by Object Class, SAG,
                  and DIR/DASA (such as CE and FOI). Functional Area detail is
                  rolled into each SAG.
                </p>
              </div>
              <div className="overflow-x-auto">
                <table className="min-w-[1550px] w-full text-sm">
                  <thead className="bg-primary text-left text-primary-foreground">
                    <tr>
                      {[
                        'DIR/DASA',
                        'Object Class',
                        'COM/OBL',
                        'SAG',
                        ...months,
                      ].map((header) => (
                        <th
                          key={header}
                          className={`whitespace-nowrap border-b border-white/10 px-3 py-3 font-semibold ${months.includes(header) ? 'text-center' : 'text-left'}`}
                        >
                          {header}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {agencyRows.map((row) => (
                      <tr
                        key={[
                          row.objectClass,
                          row.comObl,
                          row.sag,
                          row.dirDas,
                        ].join('|')}
                        className="border-b last:border-0 hover:bg-muted/30"
                      >
                        <td className="px-3 py-3">{row.dirDas}</td>
                        <td className="px-3 py-3 font-medium">
                          {row.objectClass}
                        </td>
                        <td className="px-3 py-3">{row.comObl}</td>
                        <td className="px-3 py-3 font-mono">{row.sag}</td>
                        {row.months.map((value, index) => (
                          <td
                            key={index}
                            className="px-3 py-3 text-right tabular-nums"
                          >
                            {currency(value)}
                          </td>
                        ))}
                      </tr>
                    ))}
                    <tr className="bg-[#142541] text-white">
                      <td colSpan={4} className="px-3 py-3 font-semibold">
                        {selectedSagComObl} agency total
                      </td>
                      {sagTotals.map((value, index) => (
                        <td
                          key={index}
                          className="px-3 py-3 text-right font-semibold tabular-nums"
                        >
                          {currency(value)}
                        </td>
                      ))}
                    </tr>
                    <tr className="bg-[#20395e] text-white">
                      <td colSpan={4} className="px-3 py-3 font-semibold">
                        Percent of {fiscalYear} {selectedSagComObl} total
                      </td>
                      {sagTotals.map((value, index) => (
                        <td
                          key={index}
                          className={`px-3 py-3 text-right font-semibold tabular-nums ${percentageTone(sagTotals, sagFyTotal, index)}`}
                        >
                          {percent(sagFyTotal ? value / sagFyTotal : 0)}
                          {milestoneTarget(index) !== null && (
                            <span className="ml-1 text-[10px] opacity-75">
                              /{percent(milestoneTarget(index) ?? 0)}
                            </span>
                          )}
                          {milestoneShortfall(value, sagFyTotal, index) !==
                            null && (
                            <span className="mt-0.5 block text-[10px]">
                              {currency(
                                milestoneShortfall(
                                  value,
                                  sagFyTotal,
                                  index,
                                ) ?? 0,
                              )}
                            </span>
                          )}
                        </td>
                      ))}
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
            <section
              data-print-surface="consolidated"
              className={`${activeView === 'consolidated' ? 'block' : 'hidden'} mt-6 overflow-hidden rounded-2xl border bg-card shadow-sm`}
            >
              <div className="print-report-header border-b p-5">
                <img
                  className="print-report-logo"
                  src="/esd-logo.png"
                  alt="Executive Services Directorate logo"
                />
                <div>
                  <h3 className="font-semibold">
                    {fiscalYear} Consolidated Agency Spend Plan
                  </h3>
                  <p className="text-sm text-muted-foreground">
                    Monthly values are cumulative balances through September.
                  </p>
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="min-w-[1850px] w-full text-sm">
                  <thead className="bg-muted/60 text-left">
                    <tr>
                      {[
                        'DIR/DASA',
                        'Object Class',
                        'COM/OBL',
                        'Functional Area',
                        'Status',
                        ...months,
                      ].map((header) => (
                        <th
                          key={header}
                          className={`whitespace-nowrap border-b px-3 py-3 font-semibold ${months.includes(header) ? 'text-center' : 'text-left'}`}
                        >
                          {header}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {sortedFilteredResults.map((row) => (
                      <tr
                        key={[
                          row.objectClass,
                          row.comObl,
                          row.sag,
                          row.functionalArea,
                          row.dirDas,
                        ].join('|')}
                        className="border-b last:border-0 hover:bg-muted/30"
                      >
                        <td className="px-3 py-3">{row.dirDas}</td>
                        <td className="px-3 py-3 font-medium">
                          {row.objectClass}
                        </td>
                        <td className="px-3 py-3">{row.comObl}</td>
                        <td className="px-3 py-3 font-mono text-xs">
                          {row.functionalArea}
                        </td>
                        <td className="px-3 py-3">
                          <Badge
                            variant={
                              row.status === 'Matched' ? 'default' : 'secondary'
                            }
                            title={
                              row.status === 'Multiple matches'
                                ? '“Multiple matches” means the Functional Area matched more than one DIR/DASA in the APE lookup.'
                                : undefined
                            }
                          >
                            {row.status}
                          </Badge>
                        </td>
                        {row.months.map((value, index) => (
                          <td
                            key={index}
                            className="px-3 py-3 text-right tabular-nums"
                          >
                            {currency(value)}
                          </td>
                        ))}
                      </tr>
                    ))}
                    <tr className="bg-[#142541] text-white">
                      <td colSpan={5} className="px-3 py-3 font-semibold">
                        {selectedComObl} agency total
                      </td>
                      {totals.map((value, index) => (
                        <td
                          key={index}
                          className="px-3 py-3 text-right font-semibold tabular-nums"
                        >
                          {currency(value)}
                        </td>
                      ))}
                    </tr>
                    <tr className="bg-[#20395e] text-white">
                      <td colSpan={5} className="px-3 py-3 font-semibold">
                        Percent of {fiscalYear} {selectedComObl} total
                      </td>
                      {totals.map((value, index) => (
                        <td
                          key={index}
                          className={`px-3 py-3 text-right font-semibold tabular-nums ${percentageTone(totals, fyTotal, index)}`}
                        >
                          {percent(fyTotal ? value / fyTotal : 0)}
                          {milestoneTarget(index) !== null && (
                            <span className="ml-1 text-[10px] opacity-75">
                              /{percent(milestoneTarget(index) ?? 0)}
                            </span>
                          )}
                          {milestoneShortfall(value, fyTotal, index) !==
                            null && (
                            <span className="mt-0.5 block text-[10px]">
                              {currency(
                                milestoneShortfall(value, fyTotal, index) ?? 0,
                              )}
                            </span>
                          )}
                        </td>
                      ))}
                    </tr>
                  </tbody>
                </table>
              </div>
            </section>
            <Button
              variant="outline"
              className="mt-5"
              onClick={() => {
                setSpendRows([]);
                setApeRows([]);
                setSpendName('');
                setApeName('');
                setError('');
              }}
            >
              <X /> Start over
            </Button>
          </>
        )}
      </section>
    </main>
  );
}

function UploadCard({
  title,
  description,
  name,
  ready,
  onClick,
}: {
  title: string;
  description: string;
  name: string;
  ready: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="group rounded-2xl border-2 border-dashed border-primary/25 bg-card p-6 text-left transition hover:border-primary/60 hover:bg-primary/[0.03]"
    >
      <span
        className={`mb-5 grid size-12 place-items-center rounded-xl ${ready ? 'bg-primary text-primary-foreground' : 'bg-primary/10 text-primary'}`}
      >
        {ready ? (
          <CheckCircle2 className="size-5" />
        ) : (
          <UploadCloud className="size-5" />
        )}
      </span>
      <span className="block font-semibold">{title}</span>
      <span className="mt-1 block text-sm text-muted-foreground">
        {ready ? name : description}
      </span>
      <span className="mt-5 inline-flex items-center gap-2 text-sm font-medium text-primary">
        {ready ? 'Replace workbook' : 'Choose workbook'}{' '}
        <FileSpreadsheet className="size-4" />
      </span>
    </button>
  );
}
function FundingTypeButtons({
  value,
  onChange,
}: {
  value: 'COM' | 'OBL';
  onChange: (value: 'COM' | 'OBL') => void;
}) {
  return (
    <fieldset className="grid gap-1">
      <legend className="text-sm font-medium">Funding type</legend>
      <div
        className="inline-flex rounded-md border bg-[#eef6ff] p-1"
        aria-label="Select COM or OBL"
      >
        <Button
          type="button"
          size="sm"
          variant={value === 'COM' ? 'default' : 'ghost'}
          onClick={() => onChange('COM')}
        >
          COM
        </Button>
        <Button
          type="button"
          size="sm"
          variant={value === 'OBL' ? 'default' : 'ghost'}
          onClick={() => onChange('OBL')}
        >
          OBL
        </Button>
      </div>
    </fieldset>
  );
}
function Stat({
  label,
  value,
  icon: Icon,
}: {
  label: string;
  value: string;
  icon: typeof Layers3;
}) {
  return (
    <div className="rounded-2xl border bg-card p-5 shadow-sm">
      <Icon className="mb-3 size-5 text-primary" />
      <p className="text-sm text-muted-foreground">{label}</p>
      <p className="mt-1 text-2xl font-semibold tracking-tight">{value}</p>
    </div>
  );
}

