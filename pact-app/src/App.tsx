import React from 'react';
import { HashRouter as Router, Routes, Route } from 'react-router-dom';
import { Layout } from './components/Layout';
import { DashboardPage } from './components/Dashboard/DashboardPage';
import { CasesListPage } from './components/Cases/CasesListPage';
import { NewCaseForm } from './components/Cases/NewCaseForm';
import { CaseDetail } from './components/Cases/CaseDetail';
import { StaffDirectory } from './components/Staff/StaffDirectory';
import { EscalationLog } from './components/Escalations/EscalationLog';
import { AppealsPage } from './components/Appeals/AppealsPage';
import { PolicyLibrary } from './components/Policies/PolicyLibrary';
import { StaffProfile } from './components/Staff/StaffProfile';
import { MailLogPage } from './components/Admin/MailLogPage';

// Placeholder components to prevent routing crashes
const Settings = () => <div className="glass-panel" style={{padding: '2rem'}}><h2>Settings</h2><p>Coming soon...</p></div>;

export const App: React.FC = () => {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Layout />}>
          <Route index element={<DashboardPage />} />
          <Route path="cases" element={<CasesListPage />} />
          <Route path="cases/new" element={<NewCaseForm />} />
          <Route path="cases/:id" element={<CaseDetail />} />
          <Route path="staff" element={<StaffDirectory />} />
          <Route path="staff/:id" element={<StaffProfile />} />
          <Route path="escalations" element={<EscalationLog />} />
          <Route path="appeals" element={<AppealsPage />} />
          <Route path="policies" element={<PolicyLibrary />} />
          <Route path="admin/mail-log" element={<MailLogPage />} />
          <Route path="settings" element={<Settings />} />
        </Route>
      </Routes>
    </Router>
  );
};

export default App;
