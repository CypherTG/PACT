import React from 'react';
import { HashRouter as Router, Switch, Route } from 'react-router-dom';
import { Layout } from './components/Layout';
import { DashboardPage } from './components/Dashboard/DashboardPage';
import { CasesListPage } from './components/Cases/CasesListPage';
import { NewCaseForm } from './components/Cases/NewCaseForm';
import { CaseDetail } from './components/Cases/CaseDetail';
import { StaffDirectory } from './components/Staff/StaffDirectory';
import { EscalationLog } from './components/Escalations/EscalationLog';
import { PolicyLibrary } from './components/Policies/PolicyLibrary';
import { StaffProfile } from './components/Staff/StaffProfile';
import { MailLogPage } from './components/Admin/MailLogPage';
import { ResponseLayout } from './components/Response/ResponseLayout';
import { CaseResponsePage } from './components/Response/CaseResponsePage';

// Placeholder components to prevent routing crashes
const Settings = (): React.ReactElement => (
  <div className="glass-panel" style={{padding: '2rem'}}>
    <h2>Settings</h2>
    <p>Coming soon...</p>
  </div>
);

export const App = (): React.ReactElement => {
  return (
    <Router>
      <Switch>
        {/* ─── Employee Response Routes (standalone, no sidebar) ─── */}
        <Route exact path="/case-response/:caseId/:action">
          <ResponseLayout>
            <CaseResponsePage />
          </ResponseLayout>
        </Route>

        {/* ─── Admin / Internal Routes (with sidebar) ─── */}
        <Route path="/:adminPath*">
          <Layout>
            <Switch>
              <Route exact path="/" component={DashboardPage} />
              <Route exact path="/cases" component={CasesListPage} />
              <Route exact path="/cases/new" component={NewCaseForm} />
              <Route exact path="/cases/:id" component={CaseDetail} />
              <Route exact path="/staff" component={StaffDirectory} />
              <Route exact path="/staff/:id" component={StaffProfile} />
              <Route exact path="/escalations" component={EscalationLog} />
              <Route exact path="/policies" component={PolicyLibrary} />
              <Route exact path="/admin/mail-log" component={MailLogPage} />
              <Route exact path="/settings" component={Settings} />
            </Switch>
          </Layout>
        </Route>
      </Switch>
    </Router>
  );
};

export default App;
