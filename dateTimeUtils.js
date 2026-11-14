/**
 * Return the current time in the timestamp format used by ERPNext syncs.
 * Keep this helper in one place because Apps Script files share a global scope.
 */
function getCurrentISTTimestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
}
