/**
 * Microsoft Graph API helper functions
 * All functions expect a valid access token for the appropriate scope
 */

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Generic fetch wrapper with error handling
 */
async function graphFetch(url, token, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...options.headers
    }
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(
      errorData.error?.message ||
      `Graph API error: ${response.status} ${response.statusText}`
    );
  }

  return response.json();
}

/**
 * Get SharePoint sites
 * @param {string} token - Access token with Sites.ReadWrite.All scope
 * @returns {Promise<Array>} Array of site objects
 */
export async function getSites(token) {
  const data = await graphFetch(`${GRAPH_BASE}/sites?search=*`, token);
  return data.value || [];
}

/**
 * Get lists for a specific site
 * @param {string} siteId - SharePoint site ID
 * @param {string} token - Access token with Sites.ReadWrite.All scope
 * @returns {Promise<Array>} Array of list objects
 */
export async function getLists(siteId, token) {
  const data = await graphFetch(`${GRAPH_BASE}/sites/${siteId}/lists`, token);
  return data.value || [];
}

/**
 * Get columns for a specific list
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {string} token - Access token with Sites.ReadWrite.All scope
 * @returns {Promise<Array>} Array of column objects
 */
export async function getListColumns(siteId, listId, token) {
  const data = await graphFetch(
    `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/columns`,
    token
  );
  return data.value || [];
}

/**
 * Get unread emails from inbox
 * @param {string} token - Access token with Mail.Read scope
 * @param {number} top - Number of emails to retrieve (default 20)
 * @returns {Promise<Array>} Array of email message objects
 */
export async function getUnreadEmails(token, top = 20) {
  const filter = encodeURIComponent('isRead eq false');
  const url = `${GRAPH_BASE}/me/mailFolders/inbox/messages?$filter=${filter}&$top=${top}&$orderby=receivedDateTime asc`;

  const data = await graphFetch(url, token);
  return data.value || [];
}

/**
 * Create a SharePoint list item
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {Object} fields - Field values for the new item
 * @param {string} token - Access token with Sites.ReadWrite.All scope
 * @returns {Promise<Object>} Created item object
 */
export async function createListItem(siteId, listId, fields, token) {
  const url = `${GRAPH_BASE}/sites/${siteId}/lists/${listId}/items`;

  return graphFetch(url, token, {
    method: 'POST',
    body: JSON.stringify({ fields })
  });
}

/**
 * Mark an email as read
 * @param {string} messageId - Email message ID
 * @param {string} token - Access token with Mail.ReadWrite scope
 * @returns {Promise<void>}
 */
export async function markEmailRead(messageId, token) {
  const url = `${GRAPH_BASE}/me/messages/${messageId}`;

  await fetch(url, {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ isRead: true })
  });
}

/**
 * Generate ticket number from date (YYYYMMDDHHmm format)
 * @param {string|Date} receivedDateTime - Email received date/time
 * @returns {string} Ticket number in YYYYMMDDHHmm format
 */
export function generateTicketNumber(receivedDateTime) {
  const date = new Date(receivedDateTime);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  return `${year}${month}${day}${hours}${minutes}`;
}

/**
 * Extract route from email body using regex
 * @param {string} bodyPreview - Email body preview text
 * @returns {string|null} Extracted route or null
 */
export function extractRoute(bodyPreview) {
  const match = bodyPreview.match(/route:\s*(\S+)/i);
  return match ? match[1] : null;
}
