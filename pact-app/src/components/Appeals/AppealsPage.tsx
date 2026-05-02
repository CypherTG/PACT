import React from 'react';
import { FileSearch, Filter, Plus } from 'lucide-react';
import { sharePointService } from '../../services/SharePointService';
import { AppealForm } from './AppealForm';

export const AppealsPage: React.FC = () => {
  const [appeals, setAppeals] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [showForm, setShowForm] = React.useState(false);

  const fetchAppeals = React.useCallback(async () => {
    setLoading(true);
    const data = await sharePointService.getAppeals();
    setAppeals(data);
    setLoading(false);
  }, []);

  React.useEffect(() => {
    fetchAppeals();
  }, [fetchAppeals]);

  if (showForm) {
    return (
      <AppealForm 
        onClose={() => setShowForm(false)} 
        onSuccess={() => {
          setShowForm(false);
          fetchAppeals();
        }} 
      />
    );
  }

  if (loading) {
    return (
      <div className="cases-container">
        <div className="cases-header skeleton" style={{height: '60px', borderRadius: 'var(--radii-md)'}}></div>
        <div className="cases-table-container glass-panel skeleton" style={{height: '400px', marginTop: '2rem'}}></div>
      </div>
    );
  }

  return (
    <div className="cases-container fade-in">
      <div className="cases-header">
        <h2 style={{margin: 0, display: 'flex', alignItems: 'center', gap: '8px'}}><FileSearch size={20} color="var(--secondary)"/> Appeals Register</h2>
        <div className="cases-actions">
          <button className="btn btn-secondary"><Filter size={16} /> Filter</button>
          <button className="btn btn-primary" onClick={() => setShowForm(true)}>
            <Plus size={16} /> New Appeal
          </button>
        </div>
      </div>

      <div className="cases-table-container glass-panel">
        <table className="pact-table">
            <thead>
              <tr>
                <th>Case Ref</th>
                <th>Appellant</th>
                <th>Submission Date</th>
                <th>Reviewing Officer</th>
                <th>Decision Status</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              {appeals.map(appeal => (
                <tr key={appeal.id}>
                  <td className="link-action" style={{cursor:'pointer'}}>{appeal.caseReference}</td>
                  <td>{appeal.appellant}</td>
                  <td>{new Date(appeal.appealDate).toLocaleDateString()}</td>
                  <td>{appeal.reviewingOfficer || 'Pending Assignment'}</td>
                  <td>
                    <span className={`status-badge ${appeal.decision === 'Pending' ? 'status-unpaid' : 'status-paid'}`}>
                      {appeal.decision}
                    </span>
                  </td>
                  <td>
                    <span className="link-action" style={{cursor:'pointer'}}>Review</span>
                  </td>
                </tr>
              ))}
              {appeals.length === 0 && (
                <tr>
                  <td colSpan={6} className="text-center">No appeals found.</td>
                </tr>
              )}
            </tbody>
          </table>
      </div>
    </div>
  );
};
