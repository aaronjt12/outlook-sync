import React, { useEffect, useState } from 'react';
import './MappingPanel.css';

/**
 * MappingPanel component
 * Allows users to map app fields to SharePoint columns
 */
function MappingPanel({ columns, mapping, onChange, onSave, onClose }) {
  const [localMapping, setLocalMapping] = useState(mapping || {});
  const [hasChanges, setHasChanges] = useState(false);

  // Target fields that need to be mapped
  const targetFields = [
    { key: 'ticketnumber', label: 'Ticket Number', required: true },
    { key: 'subject', label: 'Subject', required: true },
    { key: 'route', label: 'Route', required: false },
    { key: 'description', label: 'Description', required: true },
    { key: 'user', label: 'User', required: true },
    { key: 'status', label: 'Status', required: false }
  ];

  // Auto-map on mount if no mapping exists
  useEffect(() => {
    if (columns.length > 0 && Object.keys(localMapping).length === 0) {
      const autoMapping = {};

      targetFields.forEach(field => {
        // Try to find exact match by internal name
        const exactMatch = columns.find(col =>
          col.name?.toLowerCase() === field.key.toLowerCase()
        );

        if (exactMatch) {
          autoMapping[field.key] = exactMatch.name;
        } else {
          // Try to find by display name
          const displayMatch = columns.find(col =>
            col.displayName?.toLowerCase() === field.key.toLowerCase()
          );

          if (displayMatch) {
            autoMapping[field.key] = displayMatch.name;
          }
        }
      });

      if (Object.keys(autoMapping).length > 0) {
        setLocalMapping(autoMapping);
        setHasChanges(true);
      }
    }
  }, [columns]);

  const handleChange = (fieldKey, columnName) => {
    const newMapping = { ...localMapping, [fieldKey]: columnName };
    setLocalMapping(newMapping);
    setHasChanges(true);
    onChange(fieldKey, columnName);
  };

  const handleSave = () => {
    onSave(localMapping);
    setHasChanges(false);
  };

  const isValid = () => {
    return targetFields
      .filter(f => f.required)
      .every(f => localMapping[f.key]);
  };

  // Get available columns (readable, not hidden)
  const availableColumns = columns.filter(col =>
    !col.hidden && !col.readOnly && col.name !== 'ContentType'
  );

  return (
    <div className="mapping-panel-overlay" onClick={onClose}>
      <div className="mapping-panel" onClick={(e) => e.stopPropagation()}>
        <div className="mapping-panel-header">
          <h2>Field Mapping</h2>
          <button className="close-button" onClick={onClose} title="Close">
            ×
          </button>
        </div>

        <div className="mapping-panel-content">
          <p className="mapping-instructions">
            Map your app fields to SharePoint columns. Required fields are marked with *.
          </p>

          {availableColumns.length === 0 ? (
            <div className="empty-columns">
              No columns available. Please select a list first.
            </div>
          ) : (
            <div className="mapping-fields">
              {targetFields.map(field => (
                <div key={field.key} className="mapping-field-row">
                  <label className="mapping-label">
                    {field.label}
                    {field.required && <span className="required">*</span>}
                  </label>
                  <select
                    className="mapping-select"
                    value={localMapping[field.key] || ''}
                    onChange={(e) => handleChange(field.key, e.target.value)}
                  >
                    <option value="">-- Select SharePoint Column --</option>
                    {availableColumns.map(column => (
                      <option key={column.name} value={column.name}>
                        {column.displayName || column.name}
                        {column.description && ` (${column.description})`}
                      </option>
                    ))}
                  </select>
                </div>
              ))}
            </div>
          )}

          {!isValid() && (
            <div className="validation-warning">
              Please map all required fields before saving.
            </div>
          )}
        </div>

        <div className="mapping-panel-footer">
          <button
            className="button button-secondary"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            className="button button-primary"
            onClick={handleSave}
            disabled={!isValid() || !hasChanges}
          >
            Save Mapping
          </button>
        </div>
      </div>
    </div>
  );
}

export default MappingPanel;
