import { useEffect, useMemo, useState } from 'react';

const API_BASE = import.meta.env.VITE_API_BASE_URL ||
  (typeof window !== 'undefined' ? `http://${window.location.hostname}:5100` : 'http://localhost:5100');

const EMPTY_DATA = {
  turbidity: null,
  turbidity_1_hour_prior: null,
  turbidity_2_hours_prior: null,
  turbidity_3_hours_prior: null,
  current_dam_level: null,
  dam_level_1_hour_prior: null,
  dam_level_2_hours_prior: null,
  dam_level_3_hours_prior: null,
  old_res_big_tank_level: null,
  tank_a_level: null,
  tank_b_level: null,
  tank_cd_level: null,
  old_res_status: null,
  last_active_dosing: null,
  total_treatment_hours_month: null,
  current_operator: null,
  reserved_metric: null,
  target_hour: null,
  fetched_at: null,
};

export default function App() {
  const [authUser, setAuthUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState('');
  const [loginForm, setLoginForm] = useState({ username: '', password: '' });

  const [screenData, setScreenData] = useState(EMPTY_DATA);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const [chlorineDate, setChlorineDate] = useState('');
  const [lastActiveDate, setLastActiveDate] = useState('');
  const [lastActiveHour, setLastActiveHour] = useState('');

  const metrics = useMemo(() => ([
    ['Turbidity', formatMetric(screenData.turbidity, 'NTU')],
    ['Dam Level', formatMetric(screenData.current_dam_level, 'm')],
    ['Dam -1h', formatMetric(screenData.dam_level_1_hour_prior, 'm')],
    ['Dam -2h', formatMetric(screenData.dam_level_2_hours_prior, 'm')],
    ['Dam -3h', formatMetric(screenData.dam_level_3_hours_prior, 'm')],
    ['Turbidity -1h', formatMetric(screenData.turbidity_1_hour_prior, 'NTU')],
    ['Turbidity -2h', formatMetric(screenData.turbidity_2_hours_prior, 'NTU')],
    ['Turbidity -3h', formatMetric(screenData.turbidity_3_hours_prior, 'NTU')],
    ['Old Reservoir Big Tank', formatMetric(screenData.old_res_big_tank_level, '%')],
    ['Tank A', formatMetric(screenData.tank_a_level, '%')],
    ['Tank B', formatMetric(screenData.tank_b_level, '%')],
    ['Tank C/D', formatMetric(screenData.tank_cd_level, '%')],
    ['Old Reservoir Status', fallback(screenData.old_res_status)],
    ['Last Active Treatment', fallback(screenData.last_active_dosing)],
    ['Treatment Hours (Month)', fallback(screenData.total_treatment_hours_month)],
    ['Operator On Duty', fallback(screenData.current_operator)],
    ['Last Chlorine Tank Change', fallback(screenData.reserved_metric)],
    ['Target Hour', fallback(screenData.target_hour)],
  ]), [screenData]);

  const fetchMe = async () => {
    try {
      const response = await fetch(`${API_BASE}/api/auth/me`, { credentials: 'include' });
      if (!response.ok) {
        setAuthUser(null);
        return;
      }
      const payload = await response.json();
      setAuthUser(payload.user || null);
    } catch {
      setAuthUser(null);
    } finally {
      setAuthLoading(false);
    }
  };

  const fetchScreenData = async () => {
    setLoading(true);
    setError('');
    try {
      const response = await fetch(`${API_BASE}/api/screen-data/live`, { credentials: 'include' });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || 'Failed to fetch Screen Data');
      setScreenData({ ...EMPTY_DATA, ...payload });
      setChlorineDate(payload.reserved_metric ? String(payload.reserved_metric).slice(0, 10) : '');
    } catch (err) {
      setError(err.message || 'Failed to fetch Screen Data');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchMe();
  }, []);

  useEffect(() => {
    if (!authUser) return;
    fetchScreenData();
    const interval = setInterval(fetchScreenData, 15 * 60 * 1000);
    return () => clearInterval(interval);
  }, [authUser]);

  const login = async (event) => {
    event.preventDefault();
    setAuthError('');
    try {
      const response = await fetch(`${API_BASE}/api/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify(loginForm),
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || 'Login failed');
      setAuthUser(payload.user || null);
      setLoginForm({ username: '', password: '' });
    } catch (err) {
      setAuthError(err.message || 'Login failed');
    }
  };

  const logout = async () => {
    await fetch(`${API_BASE}/api/auth/logout`, { method: 'POST', credentials: 'include' });
    setAuthUser(null);
    setScreenData(EMPTY_DATA);
  };

  const saveChlorineDate = async () => {
    setError('');
    try {
      const response = await fetch(`${API_BASE}/api/screen-data/last-chlorine-tank-change`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ date: chlorineDate || null }),
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || 'Failed to update chlorine date');
      setScreenData((prev) => ({ ...prev, reserved_metric: payload.date || null }));
    } catch (err) {
      setError(err.message || 'Failed to update chlorine date');
    }
  };

  const saveLastActive = async () => {
    setError('');
    try {
      const response = await fetch(`${API_BASE}/api/screen-data/last-active-dosing`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'include',
        body: JSON.stringify({ date: lastActiveDate || null, hour: lastActiveHour || null }),
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || 'Failed to update last active treatment');
      setScreenData((prev) => ({ ...prev, last_active_dosing: payload.value || null }));
    } catch (err) {
      setError(err.message || 'Failed to update last active treatment');
    }
  };

  if (authLoading) return <div className="min-h-screen p-6">Loading authentication...</div>;

  if (!authUser) {
    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-6">
        <form onSubmit={login} className="w-full max-w-md bg-white p-6 rounded-xl shadow space-y-4">
          <h1 className="text-xl font-bold text-slate-800">Screen Data Only Login</h1>
          <input className="w-full border rounded px-3 py-2" placeholder="Username" value={loginForm.username} onChange={(e) => setLoginForm((p) => ({ ...p, username: e.target.value }))} />
          <input type="password" className="w-full border rounded px-3 py-2" placeholder="Password" value={loginForm.password} onChange={(e) => setLoginForm((p) => ({ ...p, password: e.target.value }))} />
          {authError && <p className="text-red-600 text-sm">{authError}</p>}
          <button className="w-full bg-blue-600 text-white rounded px-3 py-2 hover:bg-blue-700" type="submit">Sign In</button>
        </form>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 p-5">
      <div className="max-w-7xl mx-auto space-y-4">
        <div className="bg-white rounded-xl shadow p-4 flex flex-wrap gap-3 items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-slate-800">WATER QUALITY - SCREEN DATA ONLY</h1>
            <p className="text-sm text-slate-600">Last updated: {screenData.fetched_at ? new Date(screenData.fetched_at).toLocaleString() : '--'}</p>
          </div>
          <div className="flex gap-2">
            <button className="bg-slate-700 text-white rounded px-3 py-2" onClick={fetchScreenData} disabled={loading}>{loading ? 'Refreshing...' : 'Refresh'}</button>
            <button className="bg-red-600 text-white rounded px-3 py-2" onClick={logout}>Logout</button>
          </div>
        </div>

        {error && <div className="bg-red-100 text-red-700 p-3 rounded">{error}</div>}

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
          {metrics.map(([label, value]) => (
            <div key={label} className="bg-white rounded-lg shadow p-4">
              <div className="text-sm text-slate-500">{label}</div>
              <div className="text-xl font-semibold text-slate-900">{value}</div>
            </div>
          ))}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          <div className="bg-white rounded-xl shadow p-4 space-y-3">
            <h2 className="font-semibold text-slate-800">Update Last Chlorine Tank Change</h2>
            <input type="date" className="w-full border rounded px-3 py-2" value={chlorineDate} onChange={(e) => setChlorineDate(e.target.value)} />
            <button className="bg-blue-600 text-white rounded px-3 py-2" onClick={saveChlorineDate}>Save</button>
          </div>
          <div className="bg-white rounded-xl shadow p-4 space-y-3">
            <h2 className="font-semibold text-slate-800">Update Last Active Treatment</h2>
            <input type="date" className="w-full border rounded px-3 py-2" value={lastActiveDate} onChange={(e) => setLastActiveDate(e.target.value)} />
            <input type="number" min="1" max="12" className="w-full border rounded px-3 py-2" placeholder="Hour (1-12)" value={lastActiveHour} onChange={(e) => setLastActiveHour(e.target.value)} />
            <button className="bg-blue-600 text-white rounded px-3 py-2" onClick={saveLastActive}>Save</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function fallback(value) {
  return value === null || value === undefined || value === '' ? '--' : String(value);
}

function formatMetric(value, unit = '') {
  if (value === null || value === undefined || value === '') return '--';
  return unit ? `${value} ${unit}` : String(value);
}
