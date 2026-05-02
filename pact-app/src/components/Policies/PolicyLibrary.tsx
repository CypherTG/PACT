import React, { useEffect, useState } from 'react';
import { Search } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { PolicyOffence } from '../../config/types';

export const PolicyLibrary: React.FC = () => {
  const [policies, setPolicies] = useState<PolicyOffence[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');

  useEffect(() => {
    sharePointService.getPolicyLibrary().then(data => {
      setPolicies(data);
      setLoading(false);
    });
  }, []);

  const filteredPolicies = policies.filter(p => 
    p.offenceName.toLowerCase().includes(searchTerm.toLowerCase()) || 
    p.category.toLowerCase().includes(searchTerm.toLowerCase())
  );

  return (
    <div className="cases-container">
      <div className="cases-header">
        <div className="search-bar glass-panel">
          <Search size={18} className="search-icon" />
          <input 
            type="text" 
            placeholder="Search offences or policies..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
      </div>

      <div className="cases-table-container glass-panel">
        {loading ? (
          <div className="loading-state">Loading policy library...</div>
        ) : (
          <table className="pact-table">
            <thead>
              <tr>
                <th>Offence / Policy Name</th>
                <th>Tier Level</th>
                <th>Category</th>
                <th>Auto-Escalate?</th>
                <th>1st Offence Action</th>
                <th>2nd Offence Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredPolicies.map(p => (
                <tr key={p.id}>
                  <td className="text-white font-medium">{p.offenceName}</td>
                  <td>
                    <span className={`status-badge ${p.tier === 'Tier 1' ? 'status-paid' : p.tier === 'Tier 2' ? 'status-unpaid' : 'status-danger'}`}>
                      {p.tier}
                    </span>
                  </td>
                  <td>{p.category}</td>
                  <td>
                    {(p.escalationTrigger && p.tier === 'Tier 1')
                      ? <span className="status-badge" style={{background: 'rgba(233,69,96,0.1)', color: 'var(--primary)'}}>Yes (3 in 6m)</span>
                      : <span className="text-muted">No</span>
                    }
                  </td>
                  <td className="text-secondary">{p.firstOffenceAction}</td>
                  <td className="text-secondary">{p.secondOffenceAction}</td>
                </tr>
              ))}
              {filteredPolicies.length === 0 && (
                <tr>
                  <td colSpan={6} className="text-center">No policies found.</td>
                </tr>
              )}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
};
