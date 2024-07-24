//Optimizing allocation of resources



function optimizeResourceAllocation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Get the header row
  var headers = data[0];
  Logger.log("Headers: " + headers);
  
  // Find the index of the relevant columns
  var taskNameIndex = headers.indexOf('Task Name');
  var assignedResourceIndex = headers.indexOf('Assigned Resource');
  var priorityIndex = headers.indexOf('Priority');
  var statusIndex = headers.indexOf('Status');
  var resourceAvailabilityIndex = headers.indexOf('Resource Availability');
  
  Logger.log("Task Name Index: " + taskNameIndex);
  Logger.log("Assigned Resource Index: " + assignedResourceIndex);
  Logger.log("Priority Index: " + priorityIndex);
  Logger.log("Status Index: " + statusIndex);
  Logger.log("Resource Availability Index: " + resourceAvailabilityIndex);
  
  // Function to parse availability data
  function parseAvailability(data) {
    var availability = {};
    data.split(',').forEach(function(resource) {
      var parts = resource.split(':');
      availability[parts[0].trim()] = parseInt(parts[1].trim(), 10);
    });
    return availability;
  }
  
  // Function to serialize availability data
  function serializeAvailability(availability) {
    return Object.keys(availability).map(function(key) {
      return key + ': ' + availability[key];
    }).join(', ');
  }
  
  // Function to shuffle an array
  function shuffle(array) {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }
  
  // Array of team members
  var teamMembers = ['XuanXuan', 'Louise Hoo', 'Lee KY', 'Alice Pang', 'David Ong', 'Joe', 'Lily', 'Syakir'];
  
  // Track initial resource availability
  var initialAvailability = parseAvailability(data[1][resourceAvailabilityIndex]);
  
  // Debug: Log initial resource availability
  Logger.log("Initial resource availability: " + JSON.stringify(initialAvailability));
  
  // Iterate through rows to assign resources to tasks
  for (var i = 1; i < data.length; i++) {
    Logger.log("Row " + (i + 1) + ": " + JSON.stringify(data[i]));
    if (data[i][assignedResourceIndex] === '' && data[i][statusIndex] === 'Not started') {
      // Shuffle the list of available resources
      var shuffledTeamMembers = shuffle(teamMembers.slice());
      
      // Find the first available resource with availability > 0
      var assignedResource = shuffledTeamMembers.find(function(member) {
        return initialAvailability[member] > 0;
      });
      
      if (assignedResource) {
        // Debug: Log the task being assigned
        Logger.log("Assigning task '" + data[i][taskNameIndex] + "' to " + assignedResource);
        
        // Assign the resource
        sheet.getRange(i + 1, assignedResourceIndex + 1).setValue(assignedResource);
        
        // Check if the value was set correctly
        var assignedValue = sheet.getRange(i + 1, assignedResourceIndex + 1).getValue();
        Logger.log("Assigned Value in cell: " + assignedValue);
        
        if (assignedValue !== assignedResource) {
          Logger.log("Error: Assigned value not correctly set. Expected " + assignedResource + " but got " + assignedValue);
        }
        
        // Update resource availability
        initialAvailability[assignedResource] -= 1;
        
        // Ensure availability is not negative
        if (initialAvailability[assignedResource] < 0) {
          initialAvailability[assignedResource] = 0;
        }
        
        // Debug: Log updated resource availability
        Logger.log("Updated resource availability: " + JSON.stringify(initialAvailability));
        
        // Update the resource availability in the sheet for each row
        var updatedAvailability = serializeAvailability(initialAvailability);
        sheet.getRange(i + 1, resourceAvailabilityIndex + 1).setValue(updatedAvailability);
      } else {
        Logger.log("No available resources to assign for task '" + data[i][taskNameIndex] + "'");
      }
    }
  }
}



