import React, { useState, useEffect } from 'react';
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import './App.css';

function App() {
  const { instance, accounts } = useMsal();
  const [accessToken, setAccessToken] = useState(null);
  const [emails, setEmails] = useState([]);
  const [selectedEmails, setSelectedEmails] = useState([]);
  const [sharePointFields, setSharePointFields] = useState([]);
  const [fieldMapping, setFieldMapping] = useState({});
  const [technicians, setTechnicians] = useState([]);
  const [selectedTechnicians, setSelectedTechnicians] = useState([]);
  const [emailTechnicianMapping, setEmailTechnicianMapping] = useState({});
  const [isLoading, setIsLoading] = useState(false);

  // App fields available for mapping
  const appFields = [
    'subject',
    'description',
    'user',
    'ticketnumber',
    'assigned to'
  ];

  // Acquire access token when user is authenticated
  useEffect(() => {
    const getToken = async () => {
      if (accounts.length > 0) {
        try {
          const response = await instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
          });
          setAccessToken(response.accessToken);
          console.log('Access token acquired successfully');
        } catch (error) {
          console.error('Silent token acquisition failed:', error);
          // If silent acquisition fails, try interactive
          try {
            const response = await instance.acquireTokenRedirect(loginRequest);
          } catch (interactiveError) {
            console.error('Interactive token acquisition failed:', interactiveError);
          }
        }
      }
    };

    getToken();
  }, [accounts, instance]);

  // Login function with account picker
  const handleLogin = async () => {
    try {
      await instance.loginRedirect({
        ...loginRequest,
        prompt: "select_account" // Always show account picker
      });
    } catch (error) {
      console.error('Login failed:', error);
      alert('Login failed. Please check your Azure AD configuration and try again.');
    }
  };

  // Fetch emails from Outlook
  const fetchEmails = async () => {
    if (!accessToken) return;

    try {
      setIsLoading(true);
      const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=50&$orderby=receivedDateTime desc', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (response.ok) {
        const data = await response.json();
        setEmails(data.value);
      } else {
        console.error('Failed to fetch emails');
      }
    } catch (error) {
      console.error('Error fetching emails:', error);
    } finally {
      setIsLoading(false);
    }
  };

  // Fetch SharePoint list fields
  const fetchSharePointFields = async () => {
    if (!accessToken) return;

    try {
      const response = await fetch('https://graph.microsoft.com/v1.0/sites/root/lists/Tickets/columns', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (response.ok) {
        const data = await response.json();
        setSharePointFields(data.value);
      } else {
        console.error('Failed to fetch SharePoint fields');
      }
    } catch (error) {
      console.error('Error fetching SharePoint fields:', error);
    }
  };

  // Handle email selection
  const handleEmailSelection = (emailId) => {
    setSelectedEmails(prev => 
      prev.includes(emailId) 
        ? prev.filter(id => id !== emailId)
        : [...prev, emailId]
    );
  };

  // Handle field mapping
  const handleFieldMapping = (appField, sharePointField) => {
    setFieldMapping(prev => ({
      ...prev,
      [appField]: sharePointField
    }));
  };

  // Technician assignment component
  const TechnicianAssignment = () => {
    const handleFileUpload = (event) => {
      const file = event.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (e) => {
        const csvContent = e.target.result;
        const lines = csvContent.split('\n');
        const headers = lines[0].split(',');
        
        const parsedTechnicians = lines.slice(1).map(line => {
          const values = line.split(',');
          return {
            name: values[0]?.trim() || '',
            email: values[1]?.trim() || ''
          };
        }).filter(tech => tech.name && tech.email);

        setTechnicians(parsedTechnicians);
        setSelectedTechnicians(parsedTechnicians.map((_, index) => index));
      };
      reader.readAsText(file);
    };

    const handleTechnicianSelection = (index) => {
      setSelectedTechnicians(prev => 
        prev.includes(index) 
          ? prev.filter(i => i !== index)
          : [...prev, index]
      );
    };

    const moveTechnician = (index, direction) => {
      const newOrder = [...selectedTechnicians];
      const currentPos = newOrder.indexOf(index);
      
      if (direction === 'up' && currentPos > 0) {
        [newOrder[currentPos], newOrder[currentPos - 1]] = [newOrder[currentPos - 1], newOrder[currentPos]];
      } else if (direction === 'down' && currentPos < newOrder.length - 1) {
        [newOrder[currentPos], newOrder[currentPos + 1]] = [newOrder[currentPos + 1], newOrder[currentPos]];
      }
      
      setSelectedTechnicians(newOrder);
    };

    const handleEmailTechnicianMapping = (emailId, technicianIndex) => {
      setEmailTechnicianMapping(prev => ({
        ...prev,
        [emailId]: technicianIndex
      }));
    };

    return (
      <div style={{ margin: '20px 0', padding: '20px', border: '1px solid #ddd', borderRadius: '8px' }}>
        <h3>Technician Assignment</h3>
        
        <div style={{ marginBottom: '20px' }}>
          <input
            type="file"
            accept=".csv"
            onChange={handleFileUpload}
            style={{ marginBottom: '10px' }}
          />
          <p style={{ fontSize: '12px', color: '#666' }}>
            Upload CSV with columns: Name, Email
          </p>
        </div>

        {technicians.length > 0 && (
          <div>
            <h4>Technicians ({selectedTechnicians.length} selected)</h4>
            {technicians.map((tech, index) => (
              <div key={index} style={{ 
                display: 'flex', 
                alignItems: 'center', 
                margin: '5px 0',
                padding: '10px',
                border: selectedTechnicians.includes(index) ? '2px solid #0078d4' : '1px solid #ddd',
                borderRadius: '4px'
              }}>
                <input
                  type="checkbox"
                  checked={selectedTechnicians.includes(index)}
                  onChange={() => handleTechnicianSelection(index)}
                  style={{ marginRight: '10px' }}
                />
                <span style={{ flex: 1 }}>
                  {tech.name} ({tech.email})
                </span>
                {selectedTechnicians.includes(index) && (
                  <div>
                    <button 
                      onClick={() => moveTechnician(index, 'up')}
                      disabled={selectedTechnicians.indexOf(index) === 0}
                      style={{ marginRight: '5px' }}
                    >
                      ↑
                    </button>
                    <button 
                      onClick={() => moveTechnician(index, 'down')}
                      disabled={selectedTechnicians.indexOf(index) === selectedTechnicians.length - 1}
                    >
                      ↓
                    </button>
                  </div>
                )}
              </div>
            ))}
          </div>
        )}

        {selectedEmails.length > 0 && selectedTechnicians.length > 0 && (
          <div style={{ marginTop: '20px' }}>
            <h4>Email-Technician Assignment</h4>
            {selectedEmails.map(emailId => {
              const email = emails.find(e => e.id === emailId);
              return (
                <div key={emailId} style={{ margin: '10px 0' }}>
                  <strong>{email?.subject}</strong>
                  <select
                    value={emailTechnicianMapping[emailId] || ''}
                    onChange={(e) => handleEmailTechnicianMapping(emailId, parseInt(e.target.value))}
                    style={{ marginLeft: '10px' }}
                  >
                    <option value="">Select technician</option>
                    {selectedTechnicians.map(index => (
                      <option key={index} value={index}>
                        {technicians[index].name}
                      </option>
                    ))}
                  </select>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  };

  // Field mapping component
  const FieldMapping = () => {
    return (
      <div style={{ margin: '20px 0', padding: '20px', border: '1px solid #ddd', borderRadius: '8px' }}>
        <h3>SharePoint Field Mapping</h3>
        {appFields.map(appField => (
          <div key={appField} style={{ margin: '10px 0' }}>
            <label style={{ display: 'inline-block', width: '120px' }}>
              {appField}:
            </label>
            <select
              value={fieldMapping[appField] || ''}
              onChange={(e) => handleFieldMapping(appField, e.target.value)}
              style={{ width: '200px' }}
            >
              <option value="">Select SharePoint field</option>
              {sharePointFields.map(field => (
                <option key={field.name} value={field.name}>
                  {field.displayName || field.name}
                </option>
              ))}
            </select>
          </div>
        ))}
      </div>
    );
  };

  // Create tickets button
  const CreateTicketsButton = () => {
    const createTickets = async () => {
      if (!accessToken || selectedEmails.length === 0) return;

      try {
        setIsLoading(true);

        for (const emailId of selectedEmails) {
          const email = emails.find(e => e.id === emailId);
          if (!email) continue;

          // Generate ticket number in military time format
          const receivedDate = new Date(email.receivedDateTime);
          const ticketNumber = receivedDate.getFullYear().toString() +
            String(receivedDate.getMonth() + 1).padStart(2, '0') +
            String(receivedDate.getDate()).padStart(2, '0') +
            String(receivedDate.getHours()).padStart(2, '0') +
            String(receivedDate.getMinutes()).padStart(2, '0');

          // Get assigned technician
          const technicianIndex = emailTechnicianMapping[emailId];
          const assignedTechnician = technicianIndex !== undefined && technicians[technicianIndex] 
            ? technicians[technicianIndex].name 
            : '';

          // Build payload based on field mapping
          const payload = {};
          
          Object.keys(fieldMapping).forEach(appField => {
            const sharePointField = fieldMapping[appField];
            if (!sharePointField) return;

            switch (appField) {
              case 'subject':
                payload[sharePointField] = email.subject;
                break;
              case 'description':
                payload[sharePointField] = email.body.content;
                break;
              case 'user':
                payload[sharePointField] = email.from.emailAddress.address;
                break;
              case 'ticketnumber':
                payload[sharePointField] = ticketNumber;
                break;
              case 'assigned to':
                payload[sharePointField] = assignedTechnician;
                break;
            }
          });

          console.log('SharePoint payload:', payload);

          const response = await fetch('https://graph.microsoft.com/v1.0/sites/root/lists/Tickets/items', {
            method: 'POST',
            headers: {
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            },
            body: JSON.stringify({
              fields: payload
            })
          });

          if (response.ok) {
            console.log(`Ticket created for email: ${email.subject}`);
          } else {
            const errorData = await response.json();
            console.error(`Failed to create ticket for ${email.subject}:`, errorData);
          }
        }

        alert('Tickets created successfully!');
      } catch (error) {
        console.error('Error creating tickets:', error);
        alert('Error creating tickets. Check console for details.');
      } finally {
        setIsLoading(false);
      }
    };

    return (
      <button 
        onClick={createTickets}
        disabled={!accessToken || selectedEmails.length === 0 || isLoading}
        style={{
          padding: '10px 20px',
          backgroundColor: '#0078d4',
          color: 'white',
          border: 'none',
          borderRadius: '4px',
          cursor: 'pointer',
          fontSize: '16px'
        }}
      >
        {isLoading ? 'Creating Tickets...' : `Create ${selectedEmails.length} Ticket(s)`}
      </button>
    );
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>SWA Ticket Sync</h1>
        
        {!accessToken ? (
          <div>
            <button
              onClick={handleLogin}
              style={{
                padding: '10px 20px',
                backgroundColor: '#0078d4',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '16px'
              }}
            >
              Login with Microsoft
            </button>
          </div>
        ) : (
          <div>
            <p>Logged in successfully!</p>
            <button onClick={fetchEmails} disabled={isLoading} style={{
              padding: '10px 20px',
              backgroundColor: '#107c10',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer',
              marginRight: '10px'
            }}>
              Fetch Emails
            </button>
            <button onClick={fetchSharePointFields} disabled={isLoading} style={{
              padding: '10px 20px',
              backgroundColor: '#107c10',
              color: 'white',
              border: 'none',
              borderRadius: '4px',
              cursor: 'pointer'
            }}>
              Load SharePoint Fields
            </button>
          </div>
        )}
      </header>

      <main style={{ padding: '20px', maxWidth: '1200px', margin: '0 auto' }}>
        {accessToken && (
          <>
            {emails.length > 0 && (
              <div style={{ marginBottom: '20px' }}>
                <h2>Emails ({emails.length})</h2>
                <div style={{ maxHeight: '400px', overflowY: 'auto', border: '1px solid #ddd', padding: '10px' }}>
                  {emails.map(email => (
                    <div key={email.id} style={{
                      padding: '10px',
                      border: selectedEmails.includes(email.id) ? '2px solid #0078d4' : '1px solid #ddd',
                      margin: '5px 0',
                      borderRadius: '4px',
                      cursor: 'pointer'
                    }} onClick={() => handleEmailSelection(email.id)}>
                      <input
                        type="checkbox"
                        checked={selectedEmails.includes(email.id)}
                        onChange={() => handleEmailSelection(email.id)}
                        style={{ marginRight: '10px' }}
                      />
                      <strong>{email.subject}</strong>
                      <br />
                      <small>From: {email.from.emailAddress.address}</small>
                      <br />
                      <small>Received: {new Date(email.receivedDateTime).toLocaleString()}</small>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <TechnicianAssignment />

            {sharePointFields.length > 0 && <FieldMapping />}

            {selectedEmails.length > 0 && <CreateTicketsButton />}
          </>
        )}
      </main>
    </div>
  );
}

export default App; 