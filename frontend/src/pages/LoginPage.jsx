import { useNavigate, Link } from 'react-router-dom';
import { useEffect, useState } from 'react';
import { api, setAuth, dashboardPathForRole, getStoredUser, getToken } from '../api/client.js';

export default function LoginPage() {
  const navigate = useNavigate();

  useEffect(() => {
    const token = getToken();
    const user = getStoredUser();
    if (token && user?.role) {
      navigate(dashboardPathForRole(user.role), { replace: true });
    }
  }, [navigate]);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  async function onSubmit(e) {
    e.preventDefault();
    setError('');
    if (!email.trim() || !password) {
      setError('Isi email dan password.');
      return;
    }
    setLoading(true);
    try {
      const data = await api.login(email.trim(), password);
      setAuth(data.token, data.user);
      navigate(dashboardPathForRole(data.user.role));
    } catch (err) {
      setError(err.message || 'Login gagal.');
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="bg-login" style={{ width: '100%', minHeight: '100vh' }}>
      <div className="container">
        <div className="login-container login-split">
          <div className="login-visual">
            <div className="login-visual-header">
              
              {/* BAGIAN LOGO YANG BARU */}
              <img src="/images/logo edja.png" alt="Logo Edja Corner" style={{ width: '60px', height: '60px', borderRadius: '50%', objectFit: 'cover', marginRight: '15px' }} />
              
              <div>
                <h2 className="brand-title">Edja Corner</h2>
                <p className="brand-tagline">Sistem manajemen minimarket yang rapi dan modern.</p>
              </div>
            </div>
            <p className="login-visual-caption">
              Login untuk Owner, Admin, Gudang, atau Customer. Setelah login Anda diarahkan ke dashboard
              sesuai role.
            </p>
          </div>

          <div className="login-panel">
            <div className="login-header">
              <h1>Masuk ke akun</h1>
              <p>Gunakan email dan password yang telah dibuat.</p>
            </div>

            {error ? (
              <p style={{ color: '#dc2626', marginBottom: 12, fontSize: '0.9rem' }} role="alert">
                {error}
              </p>
            ) : null}

            <form className="login-form" onSubmit={onSubmit}>
              <div className="form-group">
                <label htmlFor="email">
                  <i className="fas fa-envelope" /> Email
                </label>
                <input
                  type="email"
                  id="email"
                  placeholder="Masukkan email Anda"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                />
              </div>

              <div className="form-group">
                <label htmlFor="password">
                  <i className="fas fa-lock" /> Password
                </label>
                <input
                  type="password"
                  id="password"
                  placeholder="Masukkan password"
                  required
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                />
              </div>

              <div className="form-options">
                <label className="checkbox-container">
                  <input type="checkbox" /> Ingat saya
                  <span className="checkmark" />
                </label>
                <Link to="/forgot-password" className="forgot-password">
                  Lupa password?
                </Link>
              </div>

              <button type="submit" className="btn-login" disabled={loading}>
                {loading ? 'Memproses...' : 'Masuk'}
              </button>
            </form>

            <p className="login-footer" style={{ marginTop: 16, textAlign: 'center' }}>
              Belum punya akun? <Link to="/register">Daftar di sini</Link>
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
