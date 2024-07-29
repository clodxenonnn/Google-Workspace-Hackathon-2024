// This is the App-Script Code for Problem Statement 1



function assignTasks() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Get the header row
  var headers = data[0];
  var taskNameIndex = headers.indexOf('Task Name');
  var assignedToIndex = headers.indexOf('Assigned To');
  var priorityIndex = headers.indexOf('Priority');
  var statusIndex = headers.indexOf('Status');
  
  // Array of team members
  var teamMembers = ['Alice Wong', 'Bob', 'Charlie', 'David Tan','Johnson', 'Jack', 'Brian Chew', 'William Tan','Richardson', 'Anne Lai', 'Gilian Yao'];
  
  // Track workload with priority weights
  var workload = {
    'Alice Wong': 0,
    'Bob': 0,
    'Charlie': 0,
    'David Tan': 0,
    'Johnson': 0,
    'Brian Chew': 0,
    'William Tan': 0,
    'Richardson': 0,
    'Anne Lai': 0,
    'Gilian Yao': 0
  };
  
  // Define priority weights
  var priorityWeights = {
    'Critical': 4,
    'High': 3,
    'Medium': 2,
    'Low': 1
  };
  
  // Calculate initial workload based on existing tasks
  for (var i = 1; i < data.length; i++) {
    var assignedTo = data[i][assignedToIndex];
    var priority = data[i][priorityIndex];
    if (assignedTo !== '' && priorityWeights[priority]) {
      workload[assignedTo] += priorityWeights[priority];
    }
  }
  
  // Function to get weighted random choice
  function weightedRandomChoice(weights) {
    var totalWeight = Object.values(weights).reduce((sum, weight) => sum + weight, 0);
    var random = Math.random() * totalWeight;
    for (var member in weights) {
      if (random < weights[member]) {
        return member;
      }
      random -= weights[member];
    }
    // Fallback to a random member in case of rounding issues
    return teamMembers[Math.floor(Math.random() * teamMembers.length)];
  }
  
  // Iterate through rows to assign new tasks
  for (var i = 1; i < data.length; i++) {
    if (data[i][assignedToIndex] === '' && data[i][statusIndex] === 'Not started') {
      // Calculate probability weights based on current workload
      var probabilityWeights = {};
      for (var member of teamMembers) {
        probabilityWeights[member] = 1 / (workload[member] + 1);
      }
      
      // Log probability weights for debugging
      Logger.log("Probability weights: " + JSON.stringify(probabilityWeights));
      
      // Find the team member with weighted random choice
      var leastBusy = weightedRandomChoice(probabilityWeights);
      
      // Assign the task
      sheet.getRange(i + 1, assignedToIndex + 1).setValue(leastBusy);
      
      // Increment workload with priority weight
      var priority = data[i][priorityIndex];
      workload[leastBusy] += priorityWeights[priority];
    }
  }
}

