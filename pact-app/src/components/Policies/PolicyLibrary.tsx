import React, { useEffect, useState, useMemo } from 'react';
import { Search, BookOpen, AlertTriangle, CheckCircle2, ShieldAlert, ArrowUpDown, ArrowUp, ArrowDown } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import type { PolicyOffence } from '../../config/types';

type SortField = 'offenceName' | 'category' | 'tier' | 'defaultPenaltyAmount';

export const PolicyLibrary: React.FC = () => {
  const [policies, setPolicies] = useState<PolicyOffence[]>([]);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('All');
  const [sortField, setSortField] = useState<SortField>('offenceName');
  const [sortAsc, setSortAsc] = useState<boolean>(true);

  useEffect(() => {
    sharePointService.getPolicyLibrary().then(data => {
      setPolicies(data);
      setLoading(false);
    });
  }, []);

  const categories = ['All', 'Conduct', 'EHSQ', 'Project Integrity', 'Strategic'];

  const handleSort = (field: SortField) => {
    if (sortField === field) {
      setSortAsc(!sortAsc);
    } else {
      setSortField(field);
      setSortAsc(true);
    }
  };

  const filteredPolicies = useMemo(() => {
    const list = policies.filter(p => {
      const matchesSearch = 
        (p.offenceName || '').toLowerCase().includes(searchTerm.toLowerCase()) || 
        (p.category || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (p.description || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (p.firstOffenceAction || '').toLowerCase().includes(searchTerm.toLowerCase());

      const matchesCategory = selectedCategory === 'All' || p.category === selectedCategory;

      return matchesSearch && matchesCategory;
    });

    return [...list].sort((a, b) => {
      let valA: any = a[sortField] ?? '';
      let valB: any = b[sortField] ?? '';

      if (typeof valA === 'string') {
        valA = valA.toLowerCase();
        valB = (valB as string).toLowerCase();
      }

      if (valA < valB) return sortAsc ? -1 : 1;
      if (valA > valB) return sortAsc ? 1 : -1;
      return 0;
    });
  }, [policies, searchTerm, selectedCategory, sortField, sortAsc]);

  const stats = useMemo(() => {
    const total = policies.length;
    const tier1 = policies.filter(p => p.tier === 'Tier 1').length;
    const tier2 = policies.filter(p => p.tier === 'Tier 2').length;
    const tier3 = policies.filter(p => p.tier === 'Tier 3').length;
    return { total, tier1, tier2, tier3 };
  }, [policies]);

  const getCategoryBadgeClass = (category: string) => {
    switch ((category || '').toLowerCase()) {
      case 'conduct': return 'category-conduct';
      case 'ehsq': return 'category-ehsq';
      case 'project integrity': return 'category-project';
      case 'strategic': return 'category-strategic';
      default: return '';
    }
  };

  const formatCurrency = (amount: number) => {
    if (!amount || amount === 0) return <span className="text-muted" style={{ fontStyle: 'italic' }}>Exempt (₦0)</span>;
    return <span style={{ fontWeight: 600, color: 'var(--text-primary)' }}>₦{amount.toLocaleString()}</span>;
  };

  const renderSortIcon = (field: SortField) => {
    if (sortField !== field) return <ArrowUpDown size={14} style={{ opacity: 0.4, marginLeft: '4px' }} />;
    return sortAsc 
      ? <ArrowUp size={14} style={{ color: 'var(--primary)', marginLeft: '4px' }} /> 
      : <ArrowDown size={14} style={{ color: 'var(--primary)', marginLeft: '4px' }} />;
  };

  return (
    <div className="cases-container">
      {/* Header Stat Cards */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px' }}>
        <div className="glass-panel" style={{ padding: '16px 20px', display: 'flex', alignItems: 'center', gap: '16px' }}>
          <div className="kpi-icon primary" style={{ background: 'rgba(0,120,212,0.1)', color: '#0078d4', padding: '12px', borderRadius: '12px' }}>
            <BookOpen size={22} />
          </div>
          <div>
            <div className="text-secondary" style={{ fontSize: '0.8rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Total Policies</div>
            <div style={{ fontSize: '1.6rem', fontWeight: 700, color: 'var(--text-primary)' }}>{stats.total}</div>
          </div>
        </div>

        <div className="glass-panel" style={{ padding: '16px 20px', display: 'flex', alignItems: 'center', gap: '16px' }}>
          <div className="kpi-icon" style={{ background: 'rgba(0,120,212,0.15)', color: '#0078d4', padding: '12px', borderRadius: '12px' }}>
            <CheckCircle2 size={22} />
          </div>
          <div>
            <div className="text-secondary" style={{ fontSize: '0.8rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Tier 1 (Minor)</div>
            <div style={{ fontSize: '1.6rem', fontWeight: 700, color: '#0078d4' }}>{stats.tier1}</div>
          </div>
        </div>

        <div className="glass-panel" style={{ padding: '16px 20px', display: 'flex', alignItems: 'center', gap: '16px' }}>
          <div className="kpi-icon" style={{ background: 'rgba(245,158,11,0.15)', color: '#d97706', padding: '12px', borderRadius: '12px' }}>
            <AlertTriangle size={22} />
          </div>
          <div>
            <div className="text-secondary" style={{ fontSize: '0.8rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Tier 2 (Moderate)</div>
            <div style={{ fontSize: '1.6rem', fontWeight: 700, color: '#d97706' }}>{stats.tier2}</div>
          </div>
        </div>

        <div className="glass-panel" style={{ padding: '16px 20px', display: 'flex', alignItems: 'center', gap: '16px' }}>
          <div className="kpi-icon" style={{ background: 'rgba(239,68,68,0.15)', color: '#ef4444', padding: '12px', borderRadius: '12px' }}>
            <ShieldAlert size={22} />
          </div>
          <div>
            <div className="text-secondary" style={{ fontSize: '0.8rem', textTransform: 'uppercase', letterSpacing: '0.05em' }}>Tier 3 (Severe)</div>
            <div style={{ fontSize: '1.6rem', fontWeight: 700, color: '#ef4444' }}>{stats.tier3}</div>
          </div>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="cases-header" style={{ flexDirection: 'row', flexWrap: 'wrap', alignItems: 'center', gap: '16px' }}>
        <div className="policy-filter-tabs">
          {categories.map(cat => {
            const count = cat === 'All' ? policies.length : policies.filter(p => p.category === cat).length;
            return (
              <button
                key={cat}
                type="button"
                className={`policy-tab-btn ${selectedCategory === cat ? 'active' : ''}`}
                onClick={() => setSelectedCategory(cat)}
              >
                {cat} ({count})
              </button>
            );
          })}
        </div>

        <div className="search-bar glass-panel" style={{ minWidth: '280px' }}>
          <Search size={18} className="search-icon" />
          <input 
            type="text" 
            placeholder="Search offences, policies, or actions..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
          />
        </div>
      </div>

      {/* Main Table */}
      <div className="cases-table-container glass-panel">
        {loading ? (
          <div className="loading-state">Loading policy library...</div>
        ) : (
          <table className="pact-table">
            <thead>
              <tr>
                <th style={{ minWidth: '220px', cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('offenceName')}>
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    Offence / Policy Name {renderSortIcon('offenceName')}
                  </div>
                </th>
                <th style={{ width: '130px', cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('category')}>
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    Category {renderSortIcon('category')}
                  </div>
                </th>
                <th style={{ width: '110px', cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('tier')}>
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    Tier Level {renderSortIcon('tier')}
                  </div>
                </th>
                <th style={{ width: '130px', cursor: 'pointer', userSelect: 'none' }} onClick={() => handleSort('defaultPenaltyAmount')}>
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    Default Penalty {renderSortIcon('defaultPenaltyAmount')}
                  </div>
                </th>
                <th style={{ width: '140px' }}>Auto-Escalate?</th>
                <th style={{ minWidth: '160px' }}>1st Offence Action</th>
                <th style={{ minWidth: '170px' }}>2nd Offence Action</th>
                <th style={{ minWidth: '180px' }}>3rd Offence Action</th>
              </tr>
            </thead>
            <tbody>
              {filteredPolicies.map((p, idx) => {
                const code = p.infractionCode || `INF-${String(idx + 1).padStart(3, '0')}`;
                return (
                  <tr key={p.id}>
                    <td>
                      <div style={{ fontWeight: 600, color: 'var(--text-primary)', fontSize: '0.95rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ 
                          background: 'rgba(0,120,212,0.1)', 
                          color: '#0078d4', 
                          border: '1px solid rgba(0,120,212,0.2)', 
                          padding: '2px 6px', 
                          borderRadius: '4px', 
                          fontSize: '0.75rem', 
                          fontFamily: 'monospace', 
                          fontWeight: 700 
                        }}>
                          {code}
                        </span>
                        <span>{p.offenceName}</span>
                      </div>
                      {p.description && p.description !== p.offenceName && (
                        <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginTop: '2px' }}>
                          {p.description}
                        </div>
                      )}
                    </td>
                    <td>
                      <span className={`category-badge ${getCategoryBadgeClass(p.category)}`}>
                        {p.category}
                      </span>
                    </td>
                    <td>
                      <span className={`tier-badge ${p.tier === 'Tier 1' ? 'tier-1' : p.tier === 'Tier 2' ? 'tier-2' : 'tier-3'}`}>
                        {p.tier}
                      </span>
                    </td>
                    <td>{formatCurrency(p.defaultPenaltyAmount)}</td>
                    <td>
                      {p.tier === 'Tier 3' ? (
                        <span className="tier-badge tier-3" style={{ fontSize: '0.7rem' }}>Immediate</span>
                      ) : p.escalationTrigger && p.tier === 'Tier 1' ? (
                        <span className="tier-badge tier-2" style={{ fontSize: '0.7rem' }}>Yes (3+ in 6m)</span>
                      ) : (
                        <span className="text-muted" style={{ fontSize: '0.85rem' }}>No</span>
                      )}
                    </td>
                    <td className="text-secondary" style={{ fontSize: '0.88rem' }}>{p.firstOffenceAction || '-'}</td>
                    <td className="text-secondary" style={{ fontSize: '0.88rem' }}>{p.secondOffenceAction || '-'}</td>
                    <td className="text-secondary" style={{ fontSize: '0.88rem' }}>{p.thirdOffenceAction || '-'}</td>
                  </tr>
                );
              })}
              {filteredPolicies.length === 0 && (
                <tr>
                  <td colSpan={8} className="text-center" style={{ padding: '30px', color: 'var(--text-muted)' }}>
                    No matching policies found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
};

