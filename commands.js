Office.onReady(() => {
  // If needed, you can add code here that runs when Office.js is ready
});

// Function called by the action button
function action(event) {
  // This function is triggered by the "Perform an action" button
  // You can add your action logic here if needed
  
  // Let the platform know we're done
  event.completed();
}

// Make sure the function is globally available
if (typeof global !== "undefined") {
  global.action = action;
}
