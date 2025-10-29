import React, { useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import EmailTable from './components/EmailTable';
import MappingPanel from './components/MappingPanel';
import { ToastContainer } from './components/Toast';
import {
  getSites,
  getLists,
  getListColumns,
  getUnreadEmails,
  createListItem,
  markEmailRead,
  generateTicketNumber,
  extractRoute
} from './lib/graph';
import {
  saveMapping,
  loadMapping,
  saveLastSite,
  loadLastSite,
  saveLastList,
  loadLastList
} from './lib/mappingStore';
import './App.css';

function App() {
  const { instance, accounts } = useMsal();

  // Dual-account tokens
  const [outlookToken, setOutlookToken] = useState(null);
  const [sharepointToken, setSharepointToken] = useState(null);
  const [outlookAccount, setOutlookAccount] = useState(null);
  const [sharepointAccount, setSharepointAccount] = useState(null);

  // SharePoint site/list state
  const [sites, setSites] = useState([]);
  const [lists, setLists] = useState([]);
  const [selectedSite, setSelectedSite] = useState(null);
  const [selectedList, setSelectedList] = useState(null);
  const [columns, setColumns] = useState([]);

  // Email state
  const [emails, setEmails] = useState([]);
  const [selectedEmailIds, setSelectedEmailIds] = useState([]);

  // Mapping state
  const [fieldMapping, setFieldMapping] = useState({});
  const [showMappingPanel, setShowMappingPanel] = useState(false);

  // UI state
  const [isLoading, setIsLoading] = useState(false);
  const [toasts, setToasts] = useState([]);

  // Load persisted site/list on mount
  useEffect(() => {
    const lastSite = loadLastSite();
    const lastList = loadLastList();
    if (lastSite) setSelectedSite(lastSite);
    if (lastList) setSelectedList(lastList);
  }, []);

  // Load mapping when site/list changes
  useEffect(() => {
    if (selectedSite && selectedList) {
      const savedMapping = loadMapping(selectedSite.id, selectedList.id);
      if (savedMapping) {
        setFieldMapping(savedMapping);
        addToast('Loaded saved field mapping', 'info', 3000);
      } else {
        setFieldMapping({});
      }
    }
  }, [selectedSite?.id, selectedList?.id]);

  // Acquire token for Outlook
  const handleOutlookLogin = async () => {
    try {
      const response = await instance.loginPopup({
        scopes: ['https://graph.microsoft.com/Mail.Read'],
        prompt: 'select_account'
      });
      setOutlookToken(response.accessToken);
      setOutlookAccount(response.account);
      addToast(`Signed in to Outlook as ${response.account.username}`, 'success');
    } catch (error) {
      console.error('Outlook login failed:', error);
      addToast('Outlook login failed: ' + error.message, 'error');
    }
  };

  // Acquire token for SharePoint
  const handleSharePointLogin = async () => {
    try {
      const response = await instance.loginPopup({
        scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All'],
        prompt: 'select_account'
      });
      setSharepointToken(response.accessToken);
      setSharepointAccount(response.account);
      addToast(`Signed in to SharePoint as ${response.account.username}`, 'success');

      // Load sites immediately after login
      loadSites(response.accessToken);
    } catch (error) {
      console.error('SharePoint login failed:', error);
      addToast('SharePoint login failed: ' + error.message, 'error');
    }
  };

  // Load SharePoint sites
  const loadSites = async (token = sharepointToken) => {
    if (!token) {
      addToast('Please sign in to SharePoint first', 'warning');
      return;
    }

    try {
      setIsLoading(true);
      const sitesData = await getSites(token);
      setSites(sitesData);
      addToast(`Loaded ${sitesData.length} sites`, 'success', 3000);
    } catch (error) {
      console.error('Failed to load sites:', error);
      addToast('Failed to load sites: ' + error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Handle site selection
  const handleSiteChange = async (siteId) => {
    const site = sites.find(s => s.id === siteId);
    if (!site) return;

    setSelectedSite({ id: site.id, name: site.displayName || site.name });
    saveLastSite(site.id, site.displayName || site.name);

    // Clear list selection
    setSelectedList(null);
    setLists([]);
    setColumns([]);

    // Load lists for this site
    try {
      setIsLoading(true);
      const listsData = await getLists(siteId, sharepointToken);
      setLists(listsData);
      addToast(`Loaded ${listsData.length} lists`, 'success', 3000);
    } catch (error) {
      console.error('Failed to load lists:', error);
      addToast('Failed to load lists: ' + error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Handle list selection
  const handleListChange = async (listId) => {
    const list = lists.find(l => l.id === listId);
    if (!list) return;

    setSelectedList({ id: list.id, name: list.displayName || list.name });
    saveLastList(list.id, list.displayName || list.name);

    // Load columns for this list
    try {
      setIsLoading(true);
      const columnsData = await getListColumns(selectedSite.id, listId, sharepointToken);
      setColumns(columnsData);
      addToast(`Loaded ${columnsData.length} columns`, 'success', 3000);
    } catch (error) {
      console.error('Failed to load columns:', error);
      addToast('Failed to load columns: ' + error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Load unread emails
  const handleLoadEmails = async () => {
    if (!outlookToken) {
      addToast('Please sign in to Outlook first', 'warning');
      return;
    }

    try {
      setIsLoading(true);
      const emailsData = await getUnreadEmails(outlookToken, 20);
      setEmails(emailsData);
      setSelectedEmailIds([]);
      addToast(`Loaded ${emailsData.length} unread emails`, 'success');
    } catch (error) {
      console.error('Failed to load emails:', error);
      addToast('Failed to load emails: ' + error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Toggle email selection
  const handleToggleEmail = (emailId) => {
    setSelectedEmailIds(prev =>
      prev.includes(emailId)
        ? prev.filter(id => id !== emailId)
        : [...prev, emailId]
    );
  };

  // Toggle all emails
  const handleToggleAllEmails = () => {
    if (selectedEmailIds.length === emails.length) {
      setSelectedEmailIds([]);
    } else {
      setSelectedEmailIds(emails.map(e => e.id));
    }
  };

  // Open mapping panel
  const handleConfigureMapping = () => {
    if (!selectedList) {
      addToast('Please select a list first', 'warning');
      return;
    }
    if (columns.length === 0) {
      addToast('No columns available for selected list', 'warning');
      return;
    }
    setShowMappingPanel(true);
  };

  // Save mapping
  const handleSaveMapping = (mapping) => {
    setFieldMapping(mapping);
    saveMapping(selectedSite.id, selectedList.id, mapping);
    setShowMappingPanel(false);
    addToast('Field mapping saved successfully', 'success');
  };

  // Check if mapping is valid
  const isMappingValid = () => {
    const required = ['ticketnumber', 'subject', 'description', 'user'];
    return required.every(field => fieldMapping[field]);
  };

  // Sync selected emails to SharePoint
  const handleSyncSelected = async () => {
    if (!outlookToken || !sharepointToken) {
      addToast('Please sign in to both Outlook and SharePoint', 'warning');
      return;
    }

    if (selectedEmailIds.length === 0) {
      addToast('Please select at least one email', 'warning');
      return;
    }

    if (!selectedSite || !selectedList) {
      addToast('Please select a site and list', 'warning');
      return;
    }

    if (!isMappingValid()) {
      addToast('Please configure field mapping first', 'warning');
      return;
    }

    try {
      setIsLoading(true);
      let successCount = 0;
      let failCount = 0;

      for (const emailId of selectedEmailIds) {
        const email = emails.find(e => e.id === emailId);
        if (!email) continue;

        try {
          // Build fields object from mapping
          const fields = {};

          // Ticket number
          if (fieldMapping.ticketnumber) {
            fields[fieldMapping.ticketnumber] = generateTicketNumber(email.receivedDateTime);
          }

          // Subject
          if (fieldMapping.subject) {
            fields[fieldMapping.subject] = email.subject || '(No Subject)';
          }

          // Route (extract from body)
          if (fieldMapping.route) {
            const route = extractRoute(email.bodyPreview);
            if (route) {
              fields[fieldMapping.route] = route;
            }
          }

          // Description
          if (fieldMapping.description) {
            fields[fieldMapping.description] = email.bodyPreview || '';
          }

          // User
          if (fieldMapping.user) {
            fields[fieldMapping.user] = email.from?.emailAddress?.address || 'Unknown';
          }

          // Status
          if (fieldMapping.status) {
            fields[fieldMapping.status] = 'New';
          }

          // Create SharePoint item
          await createListItem(selectedSite.id, selectedList.id, fields, sharepointToken);

          // Mark email as read
          await markEmailRead(emailId, outlookToken);

          successCount++;
        } catch (error) {
          console.error(`Failed to sync email ${email.subject}:`, error);
          failCount++;
        }
      }

      // Show summary
      const message = failCount > 0
        ? `Created ${successCount} tickets. ${failCount} failed.`
        : `Successfully created ${successCount} ticket(s)!`;
      addToast(message, failCount > 0 ? 'warning' : 'success');

      // Refresh emails
      await handleLoadEmails();
    } catch (error) {
      console.error('Sync failed:', error);
      addToast('Sync failed: ' + error.message, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  // Toast management
  const addToast = (message, type = 'info', duration = 5000) => {
    const id = Date.now();
    setToasts(prev => [...prev, { id, message, type, duration }]);
  };

  const removeToast = (id) => {
    setToasts(prev => prev.filter(t => t.id !== id));
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>SWA Ticket Sync</h1>
        <p className="subtitle">Sync Outlook emails to SharePoint lists</p>
      </header>

      <main className="main-content">
        {/* Auth panels */}
        <div className="auth-section">
          <div className="auth-panel">
            <h3>Outlook Connection</h3>
            {!outlookToken ? (
              <button className="button button-primary" onClick={handleOutlookLogin}>
                Sign in to Outlook
              </button>
            ) : (
              <div className="auth-status">
                <span className="status-badge status-connected">Connected</span>
                <span className="account-name">{outlookAccount?.username}</span>
              </div>
            )}
          </div>

          <div className="auth-panel">
            <h3>SharePoint Connection</h3>
            {!sharepointToken ? (
              <button className="button button-primary" onClick={handleSharePointLogin}>
                Sign in to SharePoint
              </button>
            ) : (
              <div className="auth-status">
                <span className="status-badge status-connected">Connected</span>
                <span className="account-name">{sharepointAccount?.username}</span>
              </div>
            )}
          </div>
        </div>

        {/* Site/List selection */}
        {sharepointToken && (
          <div className="selection-section">
            <div className="selection-row">
              <div className="selection-field">
                <label htmlFor="site-select">SharePoint Site</label>
                <select
                  id="site-select"
                  value={selectedSite?.id || ''}
                  onChange={(e) => handleSiteChange(e.target.value)}
                  disabled={isLoading}
                >
                  <option value="">-- Select a site --</option>
                  {sites.map(site => (
                    <option key={site.id} value={site.id}>
                      {site.displayName || site.name}
                    </option>
                  ))}
                </select>
              </div>

              <div className="selection-field">
                <label htmlFor="list-select">SharePoint List</label>
                <select
                  id="list-select"
                  value={selectedList?.id || ''}
                  onChange={(e) => handleListChange(e.target.value)}
                  disabled={!selectedSite || isLoading}
                >
                  <option value="">-- Select a list --</option>
                  {lists.map(list => (
                    <option key={list.id} value={list.id}>
                      {list.displayName || list.name}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {selectedList && (
              <div className="mapping-status">
                {isMappingValid() ? (
                  <span className="status-badge status-mapped">Field mapping configured</span>
                ) : (
                  <span className="status-badge status-warning">Field mapping required</span>
                )}
                <button
                  className="button button-secondary"
                  onClick={handleConfigureMapping}
                  disabled={isLoading}
                >
                  Configure Mapping
                </button>
              </div>
            )}
          </div>
        )}

        {/* Inbox section */}
        {outlookToken && (
          <div className="inbox-section">
            <div className="section-header">
              <h2>Inbox</h2>
              <button
                className="button button-primary"
                onClick={handleLoadEmails}
                disabled={isLoading}
              >
                {isLoading ? 'Loading...' : 'Load Unread Emails'}
              </button>
            </div>

            <EmailTable
              emails={emails}
              selectedIds={selectedEmailIds}
              onToggle={handleToggleEmail}
              onToggleAll={handleToggleAllEmails}
            />

            {selectedEmailIds.length > 0 && (
              <div className="sync-actions">
                <button
                  className="button button-sync"
                  onClick={handleSyncSelected}
                  disabled={isLoading || !isMappingValid()}
                >
                  {isLoading
                    ? 'Syncing...'
                    : `Sync ${selectedEmailIds.length} Selected Email${selectedEmailIds.length > 1 ? 's' : ''}`
                  }
                </button>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Mapping Panel Modal */}
      {showMappingPanel && (
        <MappingPanel
          columns={columns}
          mapping={fieldMapping}
          onChange={(key, value) => setFieldMapping(prev => ({ ...prev, [key]: value }))}
          onSave={handleSaveMapping}
          onClose={() => setShowMappingPanel(false)}
        />
      )}

      {/* Toast notifications */}
      <ToastContainer toasts={toasts} onRemove={removeToast} />
    </div>
  );
}

export default App;
