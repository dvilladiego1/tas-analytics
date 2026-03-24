import { useState } from 'react';
import { Calendar } from 'lucide-react';
import {
  startOfDay, endOfDay, subDays, startOfWeek, startOfMonth,
  endOfMonth, subWeeks, subMonths, format
} from 'date-fns';
import { es } from 'date-fns/locale';

const periods = [
  { id: '7d', label: '7 dias' },
  { id: 'wtd', label: 'WTD' },
  { id: 'mtd', label: 'MTD' },
  { id: 'mensual', label: 'Mensual' },
  { id: '10sem', label: '10 sem' },
  { id: 'mom', label: 'MoM' },
  { id: 'all', label: 'Todo' },
];

export function computePeriodDates(periodId, referenceDate = new Date()) {
  const now = referenceDate;
  switch (periodId) {
    case '7d':
      return { start: startOfDay(subDays(now, 6)), end: endOfDay(now) };
    case 'wtd':
      return { start: startOfWeek(now, { weekStartsOn: 1 }), end: endOfDay(now) };
    case 'mtd':
      return { start: startOfMonth(now), end: endOfDay(now) };
    case 'mensual':
      return { start: startOfMonth(now), end: endOfMonth(now) };
    case '10sem':
      return { start: startOfDay(subWeeks(now, 10)), end: endOfDay(now) };
    case 'mom': {
      return { start: startOfMonth(now), end: endOfDay(now) };
    }
    case 'all':
    default:
      return { start: null, end: null };
  }
}

export function computeComparisonDates(periodId, referenceDate = new Date()) {
  const now = referenceDate;
  switch (periodId) {
    case '7d': {
      const end = subDays(startOfDay(subDays(now, 6)), 1);
      return { start: startOfDay(subDays(end, 6)), end: endOfDay(end) };
    }
    case 'wtd': {
      const prevWeekStart = startOfWeek(subWeeks(now, 1), { weekStartsOn: 1 });
      const dayOfWeek = now.getDay() === 0 ? 6 : now.getDay() - 1;
      return { start: prevWeekStart, end: endOfDay(subDays(prevWeekStart, -dayOfWeek)) };
    }
    case 'mtd': {
      const prevMonth = subMonths(now, 1);
      const prevStart = startOfMonth(prevMonth);
      const dayOfMonth = Math.min(now.getDate(), new Date(prevMonth.getFullYear(), prevMonth.getMonth() + 1, 0).getDate());
      const prevEnd = new Date(prevMonth.getFullYear(), prevMonth.getMonth(), dayOfMonth);
      return { start: prevStart, end: endOfDay(prevEnd) };
    }
    case 'mom':
    case 'mensual': {
      const prev = subMonths(now, 1);
      return { start: startOfMonth(prev), end: endOfMonth(prev) };
    }
    case '10sem': {
      const end = subDays(startOfDay(subWeeks(now, 10)), 1);
      return { start: startOfDay(subWeeks(end, 10)), end: endOfDay(end) };
    }
    default:
      return null;
  }
}

export default function PeriodSelector({ activePeriod, onPeriodChange }) {
  const currentDates = computePeriodDates(activePeriod);
  const dateLabel = currentDates.start
    ? `${format(currentDates.start, 'dd MMM', { locale: es })} - ${format(currentDates.end, 'dd MMM yyyy', { locale: es })}`
    : 'Todos los datos';

  return (
    <div className="period-selector">
      <div className="period-buttons">
        {periods.map(({ id, label }) => (
          <button
            key={id}
            className={`period-btn ${activePeriod === id ? 'active' : ''}`}
            onClick={() => onPeriodChange(id)}
          >
            {label}
          </button>
        ))}
      </div>
      <span className="period-date-label">
        <Calendar size={13} />
        {dateLabel}
      </span>
    </div>
  );
}
