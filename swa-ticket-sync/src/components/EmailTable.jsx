import React from 'react';
import './EmailTable.css';

/**
 * EmailTable component
 * Displays a table of emails with checkboxes for selection
 */
function EmailTable({ emails, selectedIds, onToggle, onToggleAll }) {
  const allSelected = emails.length > 0 && selectedIds.length === emails.length;
  const someSelected = selectedIds.length > 0 && selectedIds.length < emails.length;

  return (
    <div className="email-table-container">
      <div className="email-table-header">
        <h3>Unread Emails ({emails.length})</h3>
        {emails.length > 0 && (
          <span className="selection-info">
            {selectedIds.length} of {emails.length} selected
          </span>
        )}
      </div>

      {emails.length === 0 ? (
        <div className="empty-state">
          No unread emails found. Load your inbox to get started.
        </div>
      ) : (
        <table className="email-table">
          <thead>
            <tr>
              <th className="checkbox-column">
                <input
                  type="checkbox"
                  checked={allSelected}
                  ref={input => {
                    if (input) {
                      input.indeterminate = someSelected;
                    }
                  }}
                  onChange={() => onToggleAll()}
                  title="Select all"
                />
              </th>
              <th className="subject-column">Subject</th>
              <th className="sender-column">Sender</th>
              <th className="received-column">Received</th>
              <th className="route-column">Route</th>
            </tr>
          </thead>
          <tbody>
            {emails.map(email => {
              const isSelected = selectedIds.includes(email.id);
              const receivedDate = new Date(email.receivedDateTime);

              // Extract route from body preview if available
              const routeMatch = email.bodyPreview?.match(/route:\s*(\S+)/i);
              const route = routeMatch ? routeMatch[1] : '—';

              return (
                <tr
                  key={email.id}
                  className={isSelected ? 'selected' : ''}
                  onClick={() => onToggle(email.id)}
                >
                  <td className="checkbox-column">
                    <input
                      type="checkbox"
                      checked={isSelected}
                      onChange={() => onToggle(email.id)}
                      onClick={(e) => e.stopPropagation()}
                    />
                  </td>
                  <td className="subject-column" title={email.subject}>
                    {email.subject || '(No Subject)'}
                  </td>
                  <td className="sender-column" title={email.from?.emailAddress?.address}>
                    {email.from?.emailAddress?.address || 'Unknown'}
                  </td>
                  <td className="received-column">
                    {receivedDate.toLocaleString()}
                  </td>
                  <td className="route-column">
                    {route}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      )}
    </div>
  );
}

export default EmailTable;
