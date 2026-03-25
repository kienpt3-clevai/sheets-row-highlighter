/* global Office */

Office.onReady(() => {
  // Commands are ready
});

// Toggle highlight from ribbon button
function toggleHighlight(event) {
  // Call the taskpane's toggle function if available
  if (window.toggleHighlight) {
    window.toggleHighlight();
  }
  event.completed();
}

// Register the function with Office
Office.actions = Office.actions || {};
Office.actions.associate("toggleHighlight", toggleHighlight);
