/**
 * LocalStorage manager for field mappings
 * Stores field mapping configurations per (siteId, listId) combination
 */

const MAPPING_KEY_PREFIX = 'fieldMapping_';
const LAST_SITE_KEY = 'lastSelectedSite';
const LAST_LIST_KEY = 'lastSelectedList';

/**
 * Generate a unique key for site+list combination
 */
function getMappingKey(siteId, listId) {
  return `${MAPPING_KEY_PREFIX}${siteId}_${listId}`;
}

/**
 * Save field mapping for a specific site/list
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @param {Object} mapping - Field mapping object
 */
export function saveMapping(siteId, listId, mapping) {
  try {
    const key = getMappingKey(siteId, listId);
    localStorage.setItem(key, JSON.stringify(mapping));
  } catch (error) {
    console.error('Failed to save mapping:', error);
  }
}

/**
 * Load field mapping for a specific site/list
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 * @returns {Object|null} Mapping object or null if not found
 */
export function loadMapping(siteId, listId) {
  try {
    const key = getMappingKey(siteId, listId);
    const stored = localStorage.getItem(key);
    return stored ? JSON.parse(stored) : null;
  } catch (error) {
    console.error('Failed to load mapping:', error);
    return null;
  }
}

/**
 * Delete field mapping for a specific site/list
 * @param {string} siteId - SharePoint site ID
 * @param {string} listId - SharePoint list ID
 */
export function deleteMapping(siteId, listId) {
  try {
    const key = getMappingKey(siteId, listId);
    localStorage.removeItem(key);
  } catch (error) {
    console.error('Failed to delete mapping:', error);
  }
}

/**
 * Save last selected site
 * @param {string} siteId - SharePoint site ID
 * @param {string} siteName - SharePoint site name (for display)
 */
export function saveLastSite(siteId, siteName) {
  try {
    localStorage.setItem(LAST_SITE_KEY, JSON.stringify({ id: siteId, name: siteName }));
  } catch (error) {
    console.error('Failed to save last site:', error);
  }
}

/**
 * Load last selected site
 * @returns {Object|null} Site object with id and name, or null
 */
export function loadLastSite() {
  try {
    const stored = localStorage.getItem(LAST_SITE_KEY);
    return stored ? JSON.parse(stored) : null;
  } catch (error) {
    console.error('Failed to load last site:', error);
    return null;
  }
}

/**
 * Save last selected list
 * @param {string} listId - SharePoint list ID
 * @param {string} listName - SharePoint list name (for display)
 */
export function saveLastList(listId, listName) {
  try {
    localStorage.setItem(LAST_LIST_KEY, JSON.stringify({ id: listId, name: listName }));
  } catch (error) {
    console.error('Failed to save last list:', error);
  }
}

/**
 * Load last selected list
 * @returns {Object|null} List object with id and name, or null
 */
export function loadLastList() {
  try {
    const stored = localStorage.getItem(LAST_LIST_KEY);
    return stored ? JSON.parse(stored) : null;
  } catch (error) {
    console.error('Failed to load last list:', error);
    return null;
  }
}

/**
 * Clear all stored mappings and selections
 */
export function clearAllMappings() {
  try {
    const keys = Object.keys(localStorage);
    keys.forEach(key => {
      if (key.startsWith(MAPPING_KEY_PREFIX) ||
          key === LAST_SITE_KEY ||
          key === LAST_LIST_KEY) {
        localStorage.removeItem(key);
      }
    });
  } catch (error) {
    console.error('Failed to clear mappings:', error);
  }
}
