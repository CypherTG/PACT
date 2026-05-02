import React, { useEffect, useState } from 'react';
import { NavLink } from 'react-router-dom';
import { Plus, Search, Filter } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { ComplianceCase } from '../../config/types';
import './Cases.css';

export const CasesListPage: React.FC = () => {
  const [cases, setCases] = useState<ComplianceCase[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');

  useEffect(() => {
    sharePointService.getCases().then(data => {
      setCases(data);
      setLoading(false);
    });
  }, []);

  const filteredCases = cases.filter(c => 
    c.title.toLowerCase().includes(searchTerm.toLowerCase()) || 
    (c.chargedPersonName || '').toLowerCase().includes(searchTerm.toLowerCase())
  );

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px', borderRadius: 'var(--radii-md)'}}></div>
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}}></div>
      </div>
    );
  }

  return (
    <div className="cases-container fade-in" data-testid="cases-list-page">
      <div className="cases-header">
        <div className="search-bar glass-panel">
          <Search size={18} className="search-icon" />
          <input 
            type="text" 
            placeholder="Search cases by ID or person..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            data-testid="cases-search-input"
          />
        </div>
        <div className="cases-actions">
          <button type="button" className="btn btn-secondary" data-testid="cases-filter-button">
            <Filter size={16} /> Filter
          </button>
          <NavLink to="/cases/new" className="btn btn-primary" data-testid="cases-new-button">
            <Plus size={16} /> New Case
          </NavLink>
        </div>
      </div>

      <div className="cases-table-container glass-panel" data-testid="cases-table">
        <table className="pact-table">
            <thead>
              <tr>
                <th>Case ID</th>
                <th>Offender</th>
                <th>Department</th>
                <th>Infraction</th>
                <th>Date</th>
                <th>Status</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredCases.map(c => (
                <tr key={c.id}>
                  <td className="font-medium text-white">{c.title}</td>
                  <td>
                    <div className="person-cell">
                      <div className="person-avatar">{c.chargedPersonName?.charAt(0) || 'U'}</div>
                      {c.chargedPersonName || c.chargedPerson}
                    </div>
                  </td>
                  <td>{c.department}</td>
                  <td>{c.offenceCategoryName || 'Unknown'}</td>
                  <td>{new Date(c.dueDate).toLocaleDateString()}</td>
                  <td>
                    <span className={`status-badge status-${c.status.toLowerCase()}`}>
                      {c.status}
                    </span>
                  </td>
                  <td>
                    <NavLink to={`/cases/${c.id}`} className="link-action">View</NavLink>
                  </td>
                </tr>
              ))}
              {filteredCases.length === 0 && (
                <tr>
                  <td colSpan={7} className="text-center">No cases found.</td>
                </tr>
              )}
            </tbody>
          </table>
      </div>
    </div>
  );
};
