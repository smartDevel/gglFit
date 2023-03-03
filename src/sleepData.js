function getSleepDataForDays(fromDaysAgo, toDaysAgo, tabName) {
    var start = new Date();
    start.setHours(0,0,0,0);
    start.setDate(start.getDate() - toDaysAgo);
  
    var end = new Date();
    end.setHours(23,59,59,999);
    end.setDate(end.getDate() - fromDaysAgo);
    
    var fitService = getFitService();
    
    var request = {
      "aggregateBy": [
        {
          "dataTypeName": "com.google.sleep.segment",
          "dataSourceId": "derived:com.google.sleep.segment:com.google.android.gms:sleep_from_activity_transition"
        }
      ],
      "bucketByTime": { "durationMillis": 86400000 },
      "startTimeMillis": start.getTime(),
      "endTimeMillis": end.getTime()
    };
    
    var response = UrlFetchApp.fetch('https://www.googleapis.com/fitness/v1/users/me/dataset:aggregate', {
      headers: {
        Authorization: 'Bearer ' + fitService.getAccessToken()
      },
      'method' : 'post',
      'contentType' : 'application/json',
      'payload' : JSON.stringify(request, null, 2)
    });
    
    var json = JSON.parse(response.getContentText());
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(tabName);
    
    for(var b = 0; b < json.bucket.length; b++) {
      // each bucket in our response should be a day
      var bucketDate = new Date(parseInt(json.bucket[b].startTimeMillis, 10));
      
      var sleepData = '';
      
      if (json.bucket[b].dataset[0].point.length > 0) {
        var sleepSegments = json.bucket[b].dataset[0].point;
        for (var i = 0; i < sleepSegments.length; i++) {
          var segment = sleepSegments[i];
          var startTime = new Date(parseInt(segment.startTimeNanos, 10) / 1000000);
          var endTime = new Date(parseInt(segment.endTimeNanos, 10) / 1000000);
          var duration = (endTime - startTime) / 1000 / 60;
          var sleepType = segment.value[0].intVal == 1 ? 'Light' : 'Deep';
          sleepData += sleepType + ' sleep: ' + duration.toFixed(1) + ' min\n';
        }
      }
      
      sheet.appendRow([bucketDate, 
                       sleepData]);
    }
  }
  