import React from 'react';
import { Search, Filter, RefreshCw } from 'lucide-react';
import { NavLink } from 'react-router-dom';
import { sharePointService } from '../../services/SharePointService';
import type { StaffMember } from '../../config/types';
import { useSharePointCollection } from '../../hooks/useSharePointCollection';

export const StaffDirectory: React.FC = () => {
  const [searchTerm, setSearchTerm] = React.useState('');
  const { data: staff, loading, refresh } = useSharePointCollection<StaffMember>(() => sharePointService.getStaffDirectory());

  const filteredStaff = staff.filter(s => 
    s.fullName.toLowerCase().includes(searchTerm.toLowerCase()) || 
    s.email.toLowerCase().includes(searchTerm.toLowerCase())
  );

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px', borderRadius: 'var(--radii-md)'}} />
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}} />
      </div>
    );
  }

  return (
    <div className="cases-container fade-in">
      <div className="cases-header">
        <div className="search-bar glass-panel">
          <Search size={18} className="search-icon" />
          <input 
            type="text" 
            placeholder="Search staff by name or email..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
        <div className="cases-actions">
          <button className="btn btn-secondary">
            <Filter size={16} /> Filter by Dept
          </button>
          <button className="btn btn-secondary" onClick={() => refresh().catch(() => undefined)}>
            <RefreshCw size={16} /> Refresh
          </button>
        </div>
      </div>

      <div className="cases-table-container glass-panel">
        <table className="pact-table">
            <thead>
              <tr>
                <th>Employee</th>
                <th>Department & Role</th>
                <th>Company</th>
                <th>Type</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredStaff.map(s => (
                <tr key={s.id}>
                  <td>
                    <div className="person-cell">
                      <div className="person-avatar" style={{ background: s.photoUrl ? 'none' : 'rgba(255,255,255,0.05)', color: 'white', overflow: 'hidden' }}>
                        {s.photoUrl ? (
                          <img src={s.photoUrl} alt="" style={{ width: '100%', height: '100%', objectFit: 'cover' }} />
                        ) : (
                          s.fullName.charAt(0)
                        )}
                      </div>
                      <div>
                        <NavLink to={`/staff/${s.id}`} className="link-action" style={{ fontWeight: 600, fontSize: '1rem', display: 'block' }}>
                          {s.fullName}
                        </NavLink>
                        <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>{s.email}</div>
                      </div>
                    </div>
                  </td>
                  <td>
                    <div>{s.department}</div>
                    <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>{s.role}</div>
                  </td>
                  <td>{s.company}</td>
                  <td>
                    <span className="status-badge" style={{background: 'rgba(255,255,255,0.05)', color: 'var(--text-secondary)'}}>
                      {s.employeeType}
                    </span>
                  </td>
                  <td>
                    <NavLink to={`/staff/${s.id}`} className="link-action">View Profile</NavLink>
                  </td>
                </tr>
              ))}
              {filteredStaff.length === 0 && (
                <tr>
                  <td colSpan={5} className="text-center">No staff members found.</td>
                </tr>
              )}
            </tbody>
          </table>
      </div>
    </div>
  );
};
