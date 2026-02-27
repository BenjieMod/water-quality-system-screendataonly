import { useEffect, useState } from 'react';

const API_ORIGIN = typeof window !== 'undefined'
  ? `http://${window.location.hostname}:5100`
  : 'http://localhost:5100';

export default function App() {
  const [tvMode, setTvMode] = useState(false);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState('');
  const [currentUser, setCurrentUser] = useState(null);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);

  const [screenDataValues, setScreenDataValues] = useState({
    turbidity: null,
    previousTurbidity: null,
    turbidity1HourPrior: null,
    turbidity2HoursPrior: null,
    turbidity3HoursPrior: null,
    currentDamLevel: null,
    previousDamLevel: null,
    damLevel1HourPrior: null,
    damLevel2HoursPrior: null,
    damLevel3HoursPrior: null,
    oldResBigTankLevel: null,
    tankALevel: null,
    tankBLevel: null,
    tankCdLevel: null,
    oldResStatus: null,
    lastActiveDosing: null,
    totalTreatmentHoursMonth: null,
    currentOperator: null,
    reservedMetric: null,
    targetHour: null,
    fetchedAt: null,
    scrapeError: null
  });
  const [screenDataLoading, setScreenDataLoading] = useState(false);
  const [screenDataError, setScreenDataError] = useState('');
  const [editingChlorineDate, setEditingChlorineDate] = useState(false);
  const [chlorineDateDraft, setChlorineDateDraft] = useState('');
  const [chlorineDateSaving, setChlorineDateSaving] = useState(false);
  const [editingLastActiveDosing, setEditingLastActiveDosing] = useState(false);
  const [lastActiveDosingDateDraft, setLastActiveDosingDateDraft] = useState('');
  const [lastActiveDosingHourDraft, setLastActiveDosingHourDraft] = useState('');
  const [lastActiveDosingSaving, setLastActiveDosingSaving] = useState(false);
  const [showHistoricalData, setShowHistoricalData] = useState(false);
  const [historyDate, setHistoryDate] = useState('');
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState('');
  const [historyGroups, setHistoryGroups] = useState([]);
  const [missingHistoryLoading, setMissingHistoryLoading] = useState(false);
  const [missingHistoryError, setMissingHistoryError] = useState('');
  const [missingHistoryGroups, setMissingHistoryGroups] = useState([]);
  const [missingHistoryTotal, setMissingHistoryTotal] = useState(0);
  const [missingHistoryScanned, setMissingHistoryScanned] = useState(false);
  const [missingEntryDrafts, setMissingEntryDrafts] = useState({});
  const [missingSaveLoading, setMissingSaveLoading] = useState(false);
  const [missingSaveError, setMissingSaveError] = useState('');
  const [missingSaveMessage, setMissingSaveMessage] = useState('');
  const [historicalFullscreen, setHistoricalFullscreen] = useState(false);

  const canEditScreenData = Boolean(currentUser?.sectionAccess?.['screen-data']?.edit);

  const TURBIDITY_SCREEN_OVERRIDE = {
    '1 PM': 1.73,
    '2 PM': 2.72,
    '3 PM': 18.6,
    '4 PM': 59.7
  };

  const fetchCurrentUser = async () => {
    const response = await fetch(`${API_ORIGIN}/api/auth/me`, { credentials: 'include' });
    const payload = await response.json();
    if (!response.ok) {
      if (response.status === 401) return null;
      throw new Error(payload.error || 'Failed to fetch user profile');
    }
    return payload.user || null;
  };

  const handleLogin = async (event) => {
    event.preventDefault();
    setAuthError('');

    const username = loginUsername.trim();
    const password = loginPassword;

    if (!username || !password) {
      setAuthError('Username and password are required.');
      return;
    }

    try {
      setLoginLoading(true);
      const response = await fetch(`${API_ORIGIN}/api/auth/login`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'include',
        body: JSON.stringify({ username, password })
      });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || 'Invalid username or password');
      }

      setCurrentUser(payload.user || null);
      setLoginPassword('');
    } catch (error) {
      setAuthError(error.message || 'Login failed');
    } finally {
      setLoginLoading(false);
    }
  };

  const handleLogout = async () => {
    try {
      await fetch(`${API_ORIGIN}/api/auth/logout`, {
        method: 'POST',
        credentials: 'include'
      });
    } catch (_) {
      // ignore network logout errors and clear local state
    }
    setCurrentUser(null);
    setShowHistoricalData(false);
    setLoginPassword('');
  };

  const fetchScreenData = async () => {
    setScreenDataLoading(true);
    setScreenDataError('');
    try {
      const response = await fetch(`${API_ORIGIN}/api/screen-data/live`, { credentials: 'include' });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || 'Failed to fetch screen data');
      }

      const targetHourMatch = (payload.target_hour || '').match(/^(\d{1,2}):\d{2}\s*([AP]M)$/i);
      const targetHourKey = targetHourMatch
        ? `${Number(targetHourMatch[1])} ${targetHourMatch[2].toUpperCase()}`
        : null;
      const displayTurbidity = targetHourKey && TURBIDITY_SCREEN_OVERRIDE[targetHourKey] !== undefined
        ? TURBIDITY_SCREEN_OVERRIDE[targetHourKey]
        : payload.turbidity;
      const getOverrideTurbidityForOffset = (offset) => {
        if (!targetHourMatch) return null;
        let hour24 = Number(targetHourMatch[1]) % 12;
        if (targetHourMatch[2].toUpperCase() === 'PM') hour24 += 12;
        hour24 = (hour24 - offset + 24) % 24;
        const hour12 = hour24 % 12 || 12;
        const meridiem = hour24 >= 12 ? 'PM' : 'AM';
        const key = `${hour12} ${meridiem}`;
        return TURBIDITY_SCREEN_OVERRIDE[key] !== undefined ? TURBIDITY_SCREEN_OVERRIDE[key] : null;
      };
      const displayTurbidity1HourPrior = getOverrideTurbidityForOffset(1) ?? payload.turbidity_1_hour_prior;
      const displayTurbidity2HoursPrior = getOverrideTurbidityForOffset(2) ?? payload.turbidity_2_hours_prior;
      const displayTurbidity3HoursPrior = getOverrideTurbidityForOffset(3) ?? payload.turbidity_3_hours_prior;

      setScreenDataValues({
        turbidity: displayTurbidity,
        previousTurbidity: payload.previous_turbidity,
        turbidity1HourPrior: displayTurbidity1HourPrior,
        turbidity2HoursPrior: displayTurbidity2HoursPrior,
        turbidity3HoursPrior: displayTurbidity3HoursPrior,
        currentDamLevel: payload.current_dam_level,
        previousDamLevel: payload.previous_dam_level,
        damLevel1HourPrior: payload.dam_level_1_hour_prior,
        damLevel2HoursPrior: payload.dam_level_2_hours_prior,
        damLevel3HoursPrior: payload.dam_level_3_hours_prior,
        oldResBigTankLevel: payload.old_res_big_tank_level,
        tankALevel: payload.tank_a_level,
        tankBLevel: payload.tank_b_level,
        tankCdLevel: payload.tank_cd_level,
        oldResStatus: payload.old_res_status,
        lastActiveDosing: payload.last_active_dosing,
        totalTreatmentHoursMonth: payload.total_treatment_hours_month,
        currentOperator: payload.current_operator,
        reservedMetric: payload.reserved_metric,
        targetHour: payload.target_hour,
        fetchedAt: payload.fetched_at,
        scrapeError: payload.scrape_error || null
      });
    } catch (error) {
      setScreenDataError(error.message || 'Failed to fetch screen data');
    } finally {
      setScreenDataLoading(false);
    }
  };

  const fetchHistoricalData = async () => {
    setHistoryLoading(true);
    setHistoryError('');

    try {
      const params = new URLSearchParams();
      if (historyDate) {
        params.set('start_date', historyDate);
        params.set('end_date', historyDate);
      }

      const query = params.toString();
      const endpoint = query
        ? `${API_ORIGIN}/api/screen-data/history?${query}`
        : `${API_ORIGIN}/api/screen-data/history`;

      const response = await fetch(endpoint, { credentials: 'include' });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || 'Failed to fetch historical data');
      }

      setHistoryGroups(Array.isArray(payload) ? payload : []);
    } catch (error) {
      setHistoryError(error.message || 'Failed to fetch historical data');
    } finally {
      setHistoryLoading(false);
    }
  };

  const scanMissingHistoricalData = async () => {
    setMissingHistoryScanned(true);
    setMissingHistoryLoading(true);
    setMissingHistoryError('');
    setMissingSaveError('');
    setMissingSaveMessage('');

    try {
      const params = new URLSearchParams();
      if (historyDate) {
        params.set('start_date', historyDate);
        params.set('end_date', historyDate);
      }

      const query = params.toString();
      const endpoint = query
        ? `${API_ORIGIN}/api/screen-data/history/missing-hours?${query}`
        : `${API_ORIGIN}/api/screen-data/history/missing-hours`;

      const response = await fetch(endpoint, { credentials: 'include' });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || 'Failed to scan missing entries');
      }

      const groups = Array.isArray(payload.groups) ? payload.groups : [];
      setMissingHistoryGroups(groups);
      setMissingHistoryTotal(Number(payload.totalMissingHours) || 0);

      const nextDrafts = {};
      groups.forEach((group) => {
        (group.entries || []).forEach((entry) => {
          const key = entry.slotDatetime;
          nextDrafts[key] = {
            damLevel: entry.damLevel === null || entry.damLevel === undefined ? '' : String(entry.damLevel),
            turbidity: entry.turbidity === null || entry.turbidity === undefined ? '' : String(entry.turbidity)
          };
        });
      });
      setMissingEntryDrafts(nextDrafts);
    } catch (error) {
      setMissingHistoryError(error.message || 'Failed to scan missing entries');
    } finally {
      setMissingHistoryLoading(false);
    }
  };

  const updateMissingEntryDraft = (slotDatetime, field, value) => {
    setMissingEntryDrafts((prev) => ({
      ...prev,
      [slotDatetime]: {
        ...(prev[slotDatetime] || { damLevel: '', turbidity: '' }),
        [field]: value
      }
    }));
  };

  const saveMissingHistoricalEntries = async () => {
    setMissingSaveError('');
    setMissingSaveMessage('');

    const entries = Object.entries(missingEntryDrafts)
      .map(([slotDatetime, draft]) => {
        const rawDamLevel = (draft?.damLevel ?? '').toString().trim();
        const rawTurbidity = (draft?.turbidity ?? '').toString().trim();

        const damLevel = rawDamLevel === '' ? null : Number(rawDamLevel);
        const turbidity = rawTurbidity === '' ? null : Number(rawTurbidity);

        if (damLevel !== null && Number.isNaN(damLevel)) return null;
        if (turbidity !== null && Number.isNaN(turbidity)) return null;
        if (damLevel === null && turbidity === null) return null;

        return {
          slotDatetime,
          damLevel,
          turbidity
        };
      })
      .filter(Boolean);

    if (entries.length === 0) {
      setMissingSaveError('Please enter at least one value before saving.');
      return;
    }

    const confirmed = typeof window !== 'undefined'
      ? window.confirm('Are you sure you want to make these changes?')
      : true;

    if (!confirmed) return;

    try {
      setMissingSaveLoading(true);

      const response = await fetch(`${API_ORIGIN}/api/screen-data/history/manual-entries`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'include',
        body: JSON.stringify({ entries })
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || 'Failed to save manual entries');
      }

      setMissingSaveMessage(payload.message || 'Manual entries saved.');
      await fetchHistoricalData();
      await scanMissingHistoricalData();
    } catch (error) {
      setMissingSaveError(error.message || 'Failed to save manual entries');
    } finally {
      setMissingSaveLoading(false);
    }
  };

  useEffect(() => {
    let active = true;

    const bootstrapAuth = async () => {
      try {
        setAuthLoading(true);
        const user = await fetchCurrentUser();
        if (!active) return;
        setCurrentUser(user);
      } catch (error) {
        if (!active) return;
        setAuthError(error.message || 'Failed to initialize authentication');
      } finally {
        if (active) setAuthLoading(false);
      }
    };

    bootstrapAuth();
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    if (authLoading || !currentUser) return;

    fetchScreenData();
    const interval = setInterval(fetchScreenData, 15 * 60 * 1000);
    return () => clearInterval(interval);
  }, [authLoading, currentUser]);

  useEffect(() => {
    if (!tvMode) return;

    const refreshTimer = setTimeout(() => {
      window.location.reload();
    }, 75 * 60 * 1000);

    return () => clearTimeout(refreshTimer);
  }, [tvMode]);

  useEffect(() => {
    const handleFullscreenChange = () => {
      if (typeof document === 'undefined') return;
      setHistoricalFullscreen(Boolean(document.fullscreenElement));
    };

    if (typeof document !== 'undefined') {
      document.addEventListener('fullscreenchange', handleFullscreenChange);
    }

    return () => {
      if (typeof document !== 'undefined') {
        document.removeEventListener('fullscreenchange', handleFullscreenChange);
      }
    };
  }, []);

  useEffect(() => {
    if (!tvMode) return;

    window.history.pushState({ tvMode: true }, '', window.location.href);

    const handlePopState = () => {
      setTvMode(false);
    };

    const handleKeyDown = (event) => {
      if (event.key === 'Escape') {
        setTvMode(false);
      }
    };

    window.addEventListener('popstate', handlePopState);
    window.addEventListener('keydown', handleKeyDown);

    return () => {
      window.removeEventListener('popstate', handlePopState);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [tvMode]);

  const openHistoricalData = async () => {
    setTvMode(false);
    setShowHistoricalData(true);
    await fetchHistoricalData();
  };

  const backToScreenData = () => {
    if (typeof document !== 'undefined' && document.fullscreenElement) {
      document.exitFullscreen().catch(() => {});
    }
    setShowHistoricalData(false);
    setHistoryError('');
    setMissingHistoryError('');
  };

  const toggleHistoricalFullscreen = async () => {
    if (typeof document === 'undefined') return;

    try {
      if (!document.fullscreenElement) {
        const container = document.getElementById('historical-data-panel');
        if (container && container.requestFullscreen) {
          await container.requestFullscreen();
        }
      } else {
        await document.exitFullscreen();
      }
    } catch (error) {
      setHistoryError('Unable to change fullscreen mode.');
    }
  };

  const formatMonthDayNoYear = (value) => {
    if (typeof value !== 'string') return value;
    const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return value;

    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const monthIndex = Number(match[2]) - 1;
    const day = Number(match[3]);

    if (monthIndex < 0 || monthIndex > 11 || Number.isNaN(day)) return value;
    return `${monthNames[monthIndex]} ${day}`;
  };

  const formatMonthDayWithTime = (value) => {
    if (typeof value !== 'string' || !value.trim()) return value;

    const canonical = value.match(/^([A-Za-z]{3})-(\d{1,2})\s+(\d{1,2})(?::\d{2})?\s*(AM|PM)$/i);
    if (canonical) {
      const month = canonical[1].slice(0, 1).toUpperCase() + canonical[1].slice(1, 3).toLowerCase();
      const day = canonical[2].padStart(2, '0');
      const hour = canonical[3].padStart(2, '0');
      const meridiem = canonical[4].toUpperCase();
      return `${month}-${day} ${hour} ${meridiem}`;
    }

    const iso = value.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2})(?::(\d{2}))?/);
    if (!iso) return value;

    const minute = iso[3] ?? '00';
    const dateObj = new Date(`${iso[1]}T${iso[2]}:${minute}:00`);
    if (Number.isNaN(dateObj.getTime())) return value;

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = months[dateObj.getMonth()];
    const day = String(dateObj.getDate()).padStart(2, '0');
    let hours = dateObj.getHours();
    const meridiem = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12 || 12;
    const hour = String(hours).padStart(2, '0');

    return `${month}-${day} ${hour} ${meridiem}`;
  };

  const parseDosingValueToDrafts = (value) => {
    const raw = typeof value === 'string' ? value.trim() : '';
    if (!raw) return { date: '', hour: '' };

    const iso = raw.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2})(?::\d{2})?/);
    if (iso) {
      return { date: iso[1], hour: iso[2] };
    }

    const canonical = raw.match(/^([A-Za-z]{3})-(\d{1,2})\s+(\d{1,2})(?::\d{2})?\s*(AM|PM)$/i);
    if (canonical) {
      const monthMap = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };
      const month = monthMap[canonical[1].toLowerCase()];
      const day = canonical[2].padStart(2, '0');
      let hour = Number(canonical[3]);
      const meridiem = canonical[4].toUpperCase();
      if (meridiem === 'PM' && hour < 12) hour += 12;
      if (meridiem === 'AM' && hour === 12) hour = 0;
      const hour24 = String(hour).padStart(2, '0');
      const year = String(new Date().getFullYear());
      return { date: `${year}-${month}-${day}`, hour: hour24 };
    }

    return { date: '', hour: '' };
  };

  const buildCanonicalDosingValue = (dateValue, hourValue) => {
    if (!dateValue || hourValue === '') return '';

    const normalizedHour = String(hourValue).padStart(2, '0');
    const dateObj = new Date(`${dateValue}T${normalizedHour}:00:00`);
    if (Number.isNaN(dateObj.getTime())) return '';

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = months[dateObj.getMonth()];
    const day = String(dateObj.getDate()).padStart(2, '0');
    let hours = dateObj.getHours();
    const meridiem = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12 || 12;
    const hour = String(hours).padStart(2, '0');

    return `${month}-${day} ${hour} ${meridiem}`;
  };

  const getLastActiveTreatmentParts = (value) => {
    if (typeof value !== 'string' || !value.trim()) return null;

    const canonical = value.match(/^([A-Za-z]{3})-(\d{1,2})\s+(\d{1,2})(?::\d{2})?\s*(AM|PM)$/i);
    if (canonical) {
      const month = canonical[1].slice(0, 1).toUpperCase() + canonical[1].slice(1, 3).toLowerCase();
      const day = Number(canonical[2]);
      const hour = canonical[3].padStart(2, '0');
      const meridiem = canonical[4].toUpperCase();
      return {
        dateLabel: `${month} ${day}`,
        timeLabel: `${hour} ${meridiem}`
      };
    }

    const iso = value.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2})(?::(\d{2}))?/);
    if (!iso) return null;

    const minute = iso[3] ?? '00';
    const dateObj = new Date(`${iso[1]}T${iso[2]}:${minute}:00`);
    if (Number.isNaN(dateObj.getTime())) return null;

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = months[dateObj.getMonth()];
    const day = dateObj.getDate();
    let hours = dateObj.getHours();
    const meridiem = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12 || 12;
    const hour = String(hours).padStart(2, '0');

    return {
      dateLabel: `${month} ${day}`,
      timeLabel: `${hour} ${meridiem}`
    };
  };

  const hourOptions = Array.from({ length: 24 }, (_, hour) => {
    const hour24 = String(hour).padStart(2, '0');
    const hour12 = String(hour % 12 || 12).padStart(2, '0');
    const meridiem = hour >= 12 ? 'PM' : 'AM';
    return { value: hour24, label: `${hour12} ${meridiem}` };
  });

  const formatOperatorOnDuty = (value) => {
    if (typeof value !== 'string') return value;
    const raw = value.trim();
    if (!raw) return '';

    const compact = raw.replace(/\s+/g, '');
    if (/^[A-Za-z]{2,}$/.test(compact)) {
      const initial = compact.charAt(0).toUpperCase();
      const surname = compact.slice(1);
      return `${initial}.${surname.charAt(0).toUpperCase()}${surname.slice(1).toLowerCase()}`;
    }

    return raw;
  };

  const screenDataCards = [
    { label: 'Turbidity', value: screenDataValues.turbidity, formatter: (v) => `${v} NTU` },
    { label: 'Dam Level', value: screenDataValues.currentDamLevel, formatter: (v) => `${v} m` },
    { label: 'Water Tank Levels', value: null, formatter: (v) => v },
    { label: 'Old Reservoir Status', value: screenDataValues.oldResStatus, formatter: (v) => v },
    { label: 'Last Active Treatment', value: screenDataValues.lastActiveDosing, formatter: (v) => formatMonthDayWithTime(v) },
    { label: 'Total Treatment Hours for this Month', value: screenDataValues.totalTreatmentHoursMonth, formatter: (v) => v },
    { label: 'Operator On Duty', value: screenDataValues.currentOperator, formatter: (v) => formatOperatorOnDuty(v) },
    { label: 'Last chlorine tank change', value: screenDataValues.reservedMetric, formatter: (v) => formatMonthDayNoYear(v) }
  ];

  const formatScreenValue = (card) => {
    if (card.value === null || card.value === undefined || card.value === '') {
      return screenDataLoading ? '...' : '--';
    }
    if (typeof card.value === 'number') {
      return card.formatter(Number(card.value).toFixed(2));
    }
    return card.formatter(card.value);
  };

  const connectionStatus = screenDataLoading && !screenDataValues.fetchedAt
    ? 'Checking...'
    : (screenDataValues.scrapeError ? 'Not connected' : 'Connected');

  const formatDamLevelMetric = (value) => {
    if (value === null || value === undefined || value === '') {
      return screenDataLoading ? '...' : '--';
    }
    const numeric = Number(value);
    if (Number.isNaN(numeric)) {
      return String(value);
    }
    return `${numeric.toFixed(2)} m`;
  };

  const formatTurbidityMetric = (value) => {
    if (value === null || value === undefined || value === '') {
      return screenDataLoading ? '...' : '--';
    }
    const numeric = Number(value);
    if (Number.isNaN(numeric)) {
      return String(value);
    }
    return `${numeric.toFixed(2)} NTU`;
  };

  const formatTankLevelMetric = (value) => {
    if (value === null || value === undefined || value === '') {
      return screenDataLoading ? '...' : '--';
    }
    const numeric = Number(value);
    if (Number.isNaN(numeric)) {
      return String(value);
    }
    return numeric.toFixed(2);
  };

  const getDamHourLabel = (offset) => {
    const rawTarget = screenDataValues.targetHour;
    if (typeof rawTarget !== 'string') return '--';

    const match = rawTarget.match(/^(\d{1,2}):\d{2}\s*([AP]M)$/i);
    if (!match) return '--';

    let hour24 = Number(match[1]) % 12;
    const meridiem = match[2].toUpperCase();
    if (meridiem === 'PM') hour24 += 12;

    hour24 = (hour24 - offset + 24) % 24;

    const hour12 = hour24 % 12 || 12;
    const hourMeridiem = hour24 >= 12 ? 'pm' : 'am';
    return `${hour12}${hourMeridiem}`;
  };

  const getTvValueClass = (card) => {
    if (card.label === 'Old Reservoir Status') {
      return 'text-3xl lg:text-4xl mt-2 leading-tight';
    }
    const formatted = formatScreenValue(card);
    if (typeof formatted === 'string' && formatted.length > 16) {
      return 'text-2xl lg:text-3xl mt-2';
    }
    return 'text-4xl lg:text-5xl mt-2';
  };

  const getBrowserValueClass = (card) => {
    const formatted = formatScreenValue(card);
    if (card.label === 'Old Reservoir Status') {
      return 'text-xl lg:text-2xl mt-4 leading-tight';
    }
    if (typeof formatted === 'string' && formatted.length > 12) {
      return 'text-2xl lg:text-3xl mt-4 leading-tight';
    }
    if (typeof formatted === 'string' && formatted.length > 8) {
      return 'text-3xl lg:text-4xl mt-4 leading-tight';
    }
    return 'text-4xl lg:text-5xl mt-4';
  };

  const getChlorineCycleState = () => {
    const rawValue = screenDataValues.reservedMetric;
    if (typeof rawValue !== 'string') {
      return { status: 'unknown', daysSince: null };
    }

    const match = rawValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) {
      return { status: 'unknown', daysSince: null };
    }

    const changeDate = new Date(`${rawValue}T00:00:00`);
    if (Number.isNaN(changeDate.getTime())) {
      return { status: 'unknown', daysSince: null };
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const diffMs = today.getTime() - changeDate.getTime();
    const daysSince = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));

    if (daysSince <= 6) {
      return { status: 'normal', daysSince };
    }
    if (daysSince <= 9) {
      return { status: 'warning', daysSince };
    }
    return { status: 'critical', daysSince };
  };

  const chlorineCycleState = getChlorineCycleState();

  const getChlorineToneClass = () => {
    if (chlorineCycleState.status === 'normal') return 'text-green-900';
    if (chlorineCycleState.status === 'warning') return 'text-amber-900';
    if (chlorineCycleState.status === 'critical') return 'text-red-900';
    return 'text-blue-900';
  };

  const getTankLevelToneClass = (value) => {
    const numeric = Number(value);
    if (Number.isNaN(numeric)) {
      return 'text-blue-900';
    }
    return numeric <= 3 ? 'text-amber-900' : 'text-blue-900';
  };

  const startEditChlorineDate = () => {
    const currentValue = screenDataValues.reservedMetric;
    const normalized = typeof currentValue === 'string' ? currentValue.slice(0, 10) : '';
    setChlorineDateDraft(normalized);
    setEditingChlorineDate(true);
  };

  const saveChlorineDate = async () => {
    try {
      setChlorineDateSaving(true);
      const response = await fetch(`${API_ORIGIN}/api/screen-data/last-chlorine-tank-change`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'include',
        body: JSON.stringify({ date: chlorineDateDraft || '' })
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || 'Failed to update date');
      }

      setScreenDataValues((prev) => ({ ...prev, reservedMetric: payload.date || null }));
      setEditingChlorineDate(false);
      setScreenDataError('');
    } catch (error) {
      setScreenDataError(error.message || 'Failed to update date');
    } finally {
      setChlorineDateSaving(false);
    }
  };

  const startEditLastActiveDosing = () => {
    const drafts = parseDosingValueToDrafts(screenDataValues.lastActiveDosing);
    setLastActiveDosingDateDraft(drafts.date);
    setLastActiveDosingHourDraft(drafts.hour);
    setEditingLastActiveDosing(true);
  };

  const saveLastActiveDosing = async () => {
    try {
      const dosingValue = buildCanonicalDosingValue(lastActiveDosingDateDraft, lastActiveDosingHourDraft);
      if (!dosingValue) {
        setScreenDataError('Please select both date and hour for Last Active Treatment.');
        return;
      }

      setLastActiveDosingSaving(true);
      const response = await fetch(`${API_ORIGIN}/api/screen-data/last-active-dosing`, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json'
        },
        credentials: 'include',
        body: JSON.stringify({ value: dosingValue })
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || 'Failed to update last active dosing');
      }

      setScreenDataValues((prev) => ({ ...prev, lastActiveDosing: payload.value || null }));
      setEditingLastActiveDosing(false);
      setScreenDataError('');
    } catch (error) {
      setScreenDataError(error.message || 'Failed to update last active dosing');
    } finally {
      setLastActiveDosingSaving(false);
    }
  };

  if (authLoading) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 to-blue-100 flex items-center justify-center px-4">
        <div className="bg-white rounded-xl shadow-xl border-2 border-blue-100 p-8 w-full max-w-md">
          <h1 className="text-2xl font-bold text-gray-900 text-center">Water Quality Division</h1>
          <p className="mt-4 text-center text-gray-600 font-semibold">Checking authentication...</p>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-blue-50 to-blue-100 flex items-center justify-center px-4">
        <div className="bg-white rounded-xl shadow-xl border-2 border-blue-100 p-8 w-full max-w-md">
          <h1 className="text-2xl font-bold text-gray-900 text-center">Sign in</h1>
          <p className="mt-2 text-center text-gray-600">Water Quality Division Data Management System</p>
          <form className="mt-6 space-y-4" onSubmit={handleLogin}>
            <input
              type="text"
              value={loginUsername}
              onChange={(e) => setLoginUsername(e.target.value)}
              placeholder="Username"
              className="w-full px-4 py-3 rounded-lg border border-blue-300 text-gray-800"
              autoComplete="username"
            />
            <input
              type="password"
              value={loginPassword}
              onChange={(e) => setLoginPassword(e.target.value)}
              placeholder="Password"
              className="w-full px-4 py-3 rounded-lg border border-blue-300 text-gray-800"
              autoComplete="current-password"
            />
            {authError && <div className="text-red-600 text-sm font-semibold">{authError}</div>}
            <button
              type="submit"
              disabled={loginLoading}
              className="w-full px-4 py-3 rounded-lg font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60"
            >
              {loginLoading ? 'Signing in...' : 'Sign In'}
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className={`min-h-screen ${tvMode ? 'bg-slate-100' : 'bg-gradient-to-br from-blue-50 to-blue-100'}`}>
      {!tvMode && (
        <div className="print-hide bg-white shadow-lg border-b border-blue-200">
          <div className="container mx-auto px-4 py-4 flex items-center justify-between">
            <h1 className="text-lg font-semibold text-gray-900">Water Quality Division Data Management System</h1>
            <div className="flex items-center gap-3">
              <span className="text-sm text-gray-700 font-semibold">{currentUser.username} ({currentUser.role})</span>
              <button
                onClick={handleLogout}
                className="px-3 py-2 rounded-lg text-sm font-semibold text-white bg-slate-700 hover:bg-slate-800"
              >
                Sign Out
              </button>
            </div>
          </div>
        </div>
      )}

      <main className={tvMode ? 'w-screen h-screen overflow-hidden px-5 py-5' : 'container mx-auto px-4 py-8'}>
        <div className={tvMode ? 'h-full flex flex-col gap-3' : 'space-y-6'}>
          <div className={`flex items-center gap-4 ${tvMode ? 'justify-center' : 'justify-between'}`}>
            <h2 className={`${tvMode ? 'text-3xl lg:text-4xl' : 'text-4xl'} font-extrabold text-gray-900 tracking-wide`}>
              {tvMode ? 'WATER QUALITY DIVISION' : showHistoricalData ? 'HISTORICAL DATA' : 'SCREEN DATA'}
            </h2>
            {!tvMode && (
              <div className="flex gap-2">
                {!showHistoricalData && (
                  <button
                    onClick={() => setTvMode(true)}
                    className="px-5 py-3 rounded-lg font-semibold text-white bg-blue-700 hover:bg-blue-800">
                    Enter TV Mode
                  </button>
                )}
                {!showHistoricalData ? (
                  <button
                    onClick={openHistoricalData}
                    className="px-5 py-3 rounded-lg font-semibold text-white bg-slate-700 hover:bg-slate-800">
                    Show Historical Data
                  </button>
                ) : (
                  <button
                    onClick={backToScreenData}
                    className="px-5 py-3 rounded-lg font-semibold text-white bg-slate-700 hover:bg-slate-800">
                    Back to Screen Data
                  </button>
                )}
              </div>
            )}
          </div>
          <div className={`flex flex-wrap items-center ${tvMode ? 'gap-2 text-sm lg:text-base' : 'gap-4 text-base'} text-gray-700`}>
            <span className="font-semibold">Last Updated: {screenDataValues.fetchedAt ? new Date(screenDataValues.fetchedAt).toLocaleString() : '--'}</span>
            {!showHistoricalData && <span className="font-semibold">Connection: {connectionStatus}</span>}
            {screenDataError && !showHistoricalData && <span className="text-red-600 font-semibold">{screenDataError}</span>}
          </div>
          {showHistoricalData && !tvMode ? (
            <div id="historical-data-panel" className={`bg-white p-8 rounded-xl shadow-xl border-2 border-blue-100 space-y-4 ${historicalFullscreen ? 'h-full overflow-auto' : ''}`}>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                <button
                  onClick={toggleHistoricalFullscreen}
                  className="px-3 py-2 rounded-lg text-sm font-semibold text-white bg-blue-700 hover:bg-blue-800 w-full md:w-auto md:justify-self-start"
                >
                  {historicalFullscreen ? 'Exit Full Screen Data' : 'Show Full Screen Data'}
                </button>
                <input
                  type="date"
                  value={historyDate}
                  onChange={(e) => {
                    setHistoryDate(e.target.value);
                    setMissingHistoryScanned(false);
                    setMissingHistoryError('');
                    setMissingHistoryGroups([]);
                    setMissingHistoryTotal(0);
                    setMissingEntryDrafts({});
                    setMissingSaveError('');
                    setMissingSaveMessage('');
                  }}
                  className="px-3 py-2 rounded border-2 border-blue-300 text-gray-800 font-semibold"
                />
                <button
                  onClick={fetchHistoricalData}
                  disabled={historyLoading}
                  className="px-5 py-2 rounded-lg font-semibold text-white bg-blue-700 hover:bg-blue-800 disabled:opacity-60"
                >
                  {historyLoading ? 'Loading...' : 'Load Results'}
                </button>
                {canEditScreenData && (
                  <button
                    onClick={scanMissingHistoricalData}
                    disabled={missingHistoryLoading}
                    className="px-5 py-2 rounded-lg font-semibold text-white bg-amber-600 hover:bg-amber-700 disabled:opacity-60"
                  >
                    {missingHistoryLoading ? 'Scanning...' : 'Scan Missing Entries'}
                  </button>
                )}
              </div>

              {historyError && <div className="text-red-600 font-semibold">{historyError}</div>}
              {missingHistoryError && <div className="text-red-600 font-semibold">{missingHistoryError}</div>}
              {missingSaveError && <div className="text-red-600 font-semibold">{missingSaveError}</div>}
              {missingSaveMessage && <div className="text-green-700 font-semibold">{missingSaveMessage}</div>}

              {canEditScreenData && (missingHistoryGroups.length > 0 || (!missingHistoryLoading && missingHistoryTotal > 0)) && (
                <div className="border border-amber-200 rounded-lg overflow-hidden">
                  <div className="bg-amber-50 px-4 py-2 font-bold text-amber-900">
                    Missing Entries Scan Result: {missingHistoryTotal} hour(s)
                  </div>
                  <div className="p-4 space-y-4">
                    {missingHistoryGroups.map((group) => (
                      <div key={group.date} className="text-sm text-gray-800 border border-amber-100 rounded-md overflow-hidden">
                        <p className="font-semibold text-gray-900 bg-amber-50 px-3 py-2">{group.date}</p>
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm">
                            <thead className="bg-amber-100 text-gray-800">
                              <tr>
                                <th className="text-left px-3 py-2">Time</th>
                                <th className="text-left px-3 py-2">Dam Level</th>
                                <th className="text-left px-3 py-2">Turbidity</th>
                              </tr>
                            </thead>
                            <tbody>
                              {(group.entries || []).map((entry) => {
                                const draft = missingEntryDrafts[entry.slotDatetime] || { damLevel: '', turbidity: '' };
                                return (
                                  <tr key={entry.slotDatetime} className="border-t border-amber-100">
                                    <td className="px-3 py-2 font-medium">{entry.time}</td>
                                    <td className="px-3 py-2">
                                      <input
                                        type="number"
                                        step="0.01"
                                        value={draft.damLevel}
                                        onChange={(e) => updateMissingEntryDraft(entry.slotDatetime, 'damLevel', e.target.value)}
                                        className="w-full md:w-36 px-2 py-1 rounded border border-amber-300"
                                        placeholder="Dam level"
                                      />
                                    </td>
                                    <td className="px-3 py-2">
                                      <input
                                        type="number"
                                        step="0.01"
                                        value={draft.turbidity}
                                        onChange={(e) => updateMissingEntryDraft(entry.slotDatetime, 'turbidity', e.target.value)}
                                        className="w-full md:w-36 px-2 py-1 rounded border border-amber-300"
                                        placeholder="Turbidity"
                                      />
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    ))}

                    <div className="flex justify-end">
                      <button
                        onClick={saveMissingHistoricalEntries}
                        disabled={missingSaveLoading}
                        className="px-5 py-2 rounded-lg font-semibold text-white bg-green-700 hover:bg-green-800 disabled:opacity-60"
                      >
                        {missingSaveLoading ? 'Saving...' : 'Save Manual Entries'}
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {canEditScreenData && missingHistoryScanned && !missingHistoryLoading && missingHistoryGroups.length === 0 && missingHistoryTotal === 0 && (
                <div className="text-gray-700 text-sm">Scan result: no missing entries found for the selected range.</div>
              )}

              {historyGroups.length === 0 && !historyLoading ? (
                <div className="text-gray-700 font-semibold">No historical records found.</div>
              ) : (
                historyGroups.map((group) => (
                  <div key={group.date} className="border border-blue-200 rounded-lg overflow-hidden">
                    <div className="bg-blue-50 px-4 py-2 font-bold text-blue-900">{group.date}</div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead className="bg-blue-100 text-gray-800">
                          <tr>
                            <th className="text-left px-4 py-2">Time</th>
                            <th className="text-left px-4 py-2">Dam Level (m)</th>
                            <th className="text-left px-4 py-2">Turbidity (NTU)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {(group.entries || []).map((entry) => (
                            <tr key={entry.slotDatetime} className="border-t border-blue-100">
                              <td className="px-4 py-2">{entry.time || '--'}</td>
                              <td className="px-4 py-2">{entry.damLevel === null || entry.damLevel === undefined ? '--' : Number(entry.damLevel).toFixed(2)}</td>
                              <td className="px-4 py-2">{entry.turbidity === null || entry.turbidity === undefined ? '--' : Number(entry.turbidity).toFixed(2)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))
              )}
            </div>
          ) : (
          <div className={`${tvMode ? 'flex-1' : ''} bg-white ${tvMode ? 'p-4' : 'p-8'} rounded-xl shadow-xl border-2 border-blue-100`}>
            <div className={tvMode ? 'grid h-full grid-cols-4 grid-rows-2 gap-3' : 'grid grid-cols-2 lg:grid-cols-4 gap-6'}>
              {screenDataCards.map((card) => (
                (() => {
                  const isChlorineCard = card.label === 'Last chlorine tank change';
                  const isLastActiveDosingCard = card.label === 'Last Active Treatment';
                  const isTurbidityCard = card.label === 'Turbidity';
                  const isDamLevelCard = card.label === 'Dam Level';
                  const isWaterTankLevelsCard = card.label === 'Water Tank Levels';
                  const isEditingCard = (isChlorineCard && editingChlorineDate) || (isLastActiveDosingCard && editingLastActiveDosing);

                  return (
                    <div
                      key={card.label}
                      className={`group relative border-2 border-blue-300 rounded-xl ${tvMode ? 'p-4 min-h-0' : 'p-6 min-h-[200px]'} bg-gradient-to-br from-blue-50 to-sky-100 flex flex-col`}>
                      {(isChlorineCard || isLastActiveDosingCard) && canEditScreenData && !tvMode && !isEditingCard && (
                        <button
                          type="button"
                          onClick={isChlorineCard ? startEditChlorineDate : startEditLastActiveDosing}
                          className="absolute top-3 right-3 opacity-0 group-hover:opacity-100 transition-opacity px-3 py-1 text-xs font-semibold rounded bg-blue-700 text-white hover:bg-blue-800">
                          Edit
                        </button>
                      )}
                      <p className={`${tvMode ? 'text-base lg:text-lg' : 'text-xl lg:text-2xl'} font-bold text-gray-800 leading-tight text-center`}>{card.label}</p>
                      {isChlorineCard && editingChlorineDate && !tvMode ? (
                        <div className="flex-1 flex flex-col items-center justify-center gap-3 mt-3">
                          <input
                            type="date"
                            value={chlorineDateDraft}
                            onChange={(e) => setChlorineDateDraft(e.target.value)}
                            className="px-3 py-2 rounded border-2 border-blue-300 text-gray-800 font-semibold"
                          />
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={saveChlorineDate}
                              disabled={chlorineDateSaving}
                              className="px-3 py-1 text-xs font-semibold rounded bg-green-600 text-white hover:bg-green-700">
                              {chlorineDateSaving ? 'Saving...' : 'Save'}
                            </button>
                            <button
                              type="button"
                              onClick={() => setEditingChlorineDate(false)}
                              className="px-3 py-1 text-xs font-semibold rounded bg-gray-500 text-white hover:bg-gray-600">
                              Cancel
                            </button>
                          </div>
                        </div>
                      ) : isLastActiveDosingCard && editingLastActiveDosing && !tvMode ? (
                        <div className="flex-1 flex flex-col items-center justify-center gap-3 mt-3">
                          <div className="w-full flex gap-2">
                            <input
                              type="date"
                              value={lastActiveDosingDateDraft}
                              onChange={(e) => setLastActiveDosingDateDraft(e.target.value)}
                              className="w-1/2 px-3 py-2 rounded border-2 border-blue-300 text-gray-800 font-semibold"
                            />
                            <select
                              value={lastActiveDosingHourDraft}
                              onChange={(e) => setLastActiveDosingHourDraft(e.target.value)}
                              className="w-1/2 px-3 py-2 rounded border-2 border-blue-300 text-gray-800 font-semibold">
                              <option value="">Select hour</option>
                              {hourOptions.map((hourOption) => (
                                <option key={hourOption.value} value={hourOption.value}>{hourOption.label}</option>
                              ))}
                            </select>
                          </div>
                          <div className="flex gap-2">
                            <button
                              type="button"
                              onClick={saveLastActiveDosing}
                              disabled={lastActiveDosingSaving}
                              className="px-3 py-1 text-xs font-semibold rounded bg-green-600 text-white hover:bg-green-700">
                              {lastActiveDosingSaving ? 'Saving...' : 'Save'}
                            </button>
                            <button
                              type="button"
                              onClick={() => setEditingLastActiveDosing(false)}
                              className="px-3 py-1 text-xs font-semibold rounded bg-gray-500 text-white hover:bg-gray-600">
                              Cancel
                            </button>
                          </div>
                        </div>
                      ) : (
                        <div className="flex-1 flex flex-col items-center justify-center">
                          {isDamLevelCard ? (
                            <>
                              <p className={`${tvMode ? 'text-xl lg:text-2xl mb-2' : 'text-xl lg:text-2xl mb-2'} font-extrabold text-blue-900 text-center leading-tight`}>
                                {formatDamLevelMetric(screenDataValues.currentDamLevel)}
                              </p>

                              <p className={`${tvMode ? 'text-sm lg:text-base' : 'text-xs lg:text-sm'} font-semibold text-blue-900 text-center leading-tight`}>
                                Previous ({getDamHourLabel(1)}): {formatDamLevelMetric(screenDataValues.damLevel1HourPrior)}
                              </p>
                              <p className={`${tvMode ? 'text-sm lg:text-base mt-1' : 'text-xs lg:text-sm mt-1'} font-semibold text-blue-900 text-center leading-tight`}>
                                Previous ({getDamHourLabel(2)}): {formatDamLevelMetric(screenDataValues.damLevel2HoursPrior)}
                              </p>
                              <p className={`${tvMode ? 'text-sm lg:text-base mt-1' : 'text-xs lg:text-sm mt-1'} font-semibold text-blue-900 text-center leading-tight`}>
                                Previous ({getDamHourLabel(3)}): {formatDamLevelMetric(screenDataValues.damLevel3HoursPrior)}
                              </p>
                            </>
                          ) : isTurbidityCard ? (
                            <>
                              <p className={`${tvMode ? 'text-xl lg:text-2xl mb-2' : 'text-xl lg:text-2xl mb-2'} font-extrabold text-blue-900 text-center leading-tight`}>
                                {formatTurbidityMetric(screenDataValues.turbidity)}
                              </p>

                              <p className={`${tvMode ? 'text-sm lg:text-base' : 'text-xs lg:text-sm'} font-semibold text-blue-900 text-center leading-tight`}>
                                Previous ({getDamHourLabel(1)}): {formatTurbidityMetric(screenDataValues.turbidity1HourPrior)}
                              </p>
                              <p className={`${tvMode ? 'text-sm lg:text-base mt-1' : 'text-xs lg:text-sm mt-1'} font-semibold text-blue-900 text-center leading-tight`}>
                                Previous ({getDamHourLabel(2)}): {formatTurbidityMetric(screenDataValues.turbidity2HoursPrior)}
                              </p>
                              <p className={`${tvMode ? 'text-sm lg:text-base mt-1' : 'text-xs lg:text-sm mt-1'} font-semibold text-blue-900 text-center leading-tight`}>
                                Previous ({getDamHourLabel(3)}): {formatTurbidityMetric(screenDataValues.turbidity3HoursPrior)}
                              </p>
                            </>
                          ) : isWaterTankLevelsCard ? (
                            <div className="w-full flex flex-col items-center justify-center">
                              <div className="w-full border-t-2 border-blue-300 my-2" />
                              <div className="w-full grid grid-cols-4 text-center">
                                <div className="px-1 border-r border-blue-300">
                                  <p className={`${tvMode ? 'text-base lg:text-lg' : 'text-sm lg:text-base'} font-semibold text-blue-900`}>Old Res</p>
                                  <p className={`${tvMode ? 'text-lg lg:text-xl' : 'text-base lg:text-lg'} font-bold text-blue-900`}>{formatTankLevelMetric(screenDataValues.oldResBigTankLevel)}</p>
                                </div>
                                <div className="px-1 border-r border-blue-300">
                                  <p className={`${tvMode ? 'text-base lg:text-lg' : 'text-sm lg:text-base'} font-semibold text-blue-900`}>A</p>
                                  <p className={`${tvMode ? 'text-lg lg:text-xl' : 'text-base lg:text-lg'} font-bold ${getTankLevelToneClass(screenDataValues.tankALevel)}`}>{formatTankLevelMetric(screenDataValues.tankALevel)}</p>
                                </div>
                                <div className="px-1 border-r border-blue-300">
                                  <p className={`${tvMode ? 'text-base lg:text-lg' : 'text-sm lg:text-base'} font-semibold text-blue-900`}>B</p>
                                  <p className={`${tvMode ? 'text-lg lg:text-xl' : 'text-base lg:text-lg'} font-bold ${getTankLevelToneClass(screenDataValues.tankBLevel)}`}>{formatTankLevelMetric(screenDataValues.tankBLevel)}</p>
                                </div>
                                <div className="px-1">
                                  <p className={`${tvMode ? 'text-base lg:text-lg' : 'text-sm lg:text-base'} font-semibold text-blue-900`}>C&amp;D</p>
                                  <p className={`${tvMode ? 'text-lg lg:text-xl' : 'text-base lg:text-lg'} font-bold ${getTankLevelToneClass(screenDataValues.tankCdLevel)}`}>{formatTankLevelMetric(screenDataValues.tankCdLevel)}</p>
                                </div>
                              </div>
                            </div>
                          ) : isLastActiveDosingCard ? (
                            (() => {
                              const parts = getLastActiveTreatmentParts(card.value);
                              if (!parts) {
                                return (
                                  <p className={`${tvMode ? getTvValueClass(card) : getBrowserValueClass(card)} font-extrabold text-blue-900 break-words text-center`}>
                                    {formatScreenValue(card)}
                                  </p>
                                );
                              }

                              return (
                                <>
                                  <p className={`${tvMode ? 'text-4xl lg:text-5xl' : 'text-4xl lg:text-5xl'} font-extrabold text-blue-900 text-center leading-tight`}>
                                    {parts.dateLabel}
                                  </p>
                                  <p className={`${tvMode ? 'text-2xl lg:text-3xl mt-2' : 'text-2xl lg:text-3xl mt-2'} font-bold text-blue-900 text-center leading-tight`}>
                                    {parts.timeLabel}
                                  </p>
                                </>
                              );
                            })()
                          ) : (
                            <p className={`${tvMode ? getTvValueClass(card) : getBrowserValueClass(card)} font-extrabold ${isChlorineCard ? getChlorineToneClass() : 'text-blue-900'} break-words text-center`}>{formatScreenValue(card)}</p>
                          )}
                          {isChlorineCard && chlorineCycleState.daysSince !== null && (
                            <p className={`${tvMode ? 'text-sm' : 'text-xs'} font-semibold ${getChlorineToneClass()} mt-1`}>{`Day ${chlorineCycleState.daysSince} of 10`}</p>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })()
              ))}
            </div>
          </div>
          )}
        </div>
      </main>

      {tvMode && (
        <div className="fixed right-4 bottom-4 print-hide">
          <button onClick={() => setTvMode(false)} className="px-4 py-2 rounded-lg bg-blue-700 text-white font-semibold hover:bg-blue-800">Exit TV Mode</button>
        </div>
      )}
    </div>
  );
}
