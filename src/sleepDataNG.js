function getSleepDataNG(fromDaysAgo, toDaysAgo, tabName) {
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
          "dataTypeName": "com.google.activity.segment",
          "dataSourceId": "derived:com.google.activity.segment:com.google.android.gms:merge_activity_segments"
        }
      ],
      "bucketByActivityType": {
        "minDurationMillis": 1,
        "activityDataSourceId": "derived:com.google.activity.segment:com.google.android.gms:merge_activity_segments",
        "activityType": "72"
      },
      "startTimeMillis": start.getTime(),
      "endTimeMillis": end.getTime()
    };
    Logger.log('Start: ' + start);
    Logger.log('End: ' + end);
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
    /*
    for(var b = 0; b < json.bucket.length; b++) {
      // each bucket in our response should be a sleep session
      var bucketStart = new Date(parseInt(json.bucket[b].startTimeMillis, 10));
      var bucketEnd = new Date(parseInt(json.bucket[b].endTimeMillis, 10));
      var durationMillis = json.bucket[b].dataset[0].point[0].endTimeNanos - json.bucket[b].dataset[0].point[0].startTimeNanos;
      var durationMinutes = durationMillis / (1000 * 60);
      
      sheet.appendRow([bucketStart, bucketEnd, durationMinutes]);
    }*/

    for(var b = 0; b < json.bucket.length; b++) {
        // each bucket in our response should be a day
        var bucketDate = new Date(parseInt(json.bucket[b].startTimeMillis, 10));
        Logger.log('BucketDate: ' + bucketDate);
        
        var sleepData = '';
        
        if (json.bucket[b].dataset[0].point.length > 0) {
          var sleepSegments = json.bucket[b].dataset[0].point;
          Logger.log('SleepSegment: ' + sleepSegments);
          for (var i = 0; i < sleepSegments.length; i++) {
            var segment = sleepSegments[i];
            

            var startTime = new Date(parseInt(segment.startTimeNanos, 10) / 1000000);
            var endTime = new Date(parseInt(segment.endTimeNanos, 10) / 1000000);
            Logger.log('Segment-Starttime' + startTime);
            Logger.log('Segment-Endtime' + endTime);
            var durationInMs = endTime - startTime;
            var duration = Math.round(durationInMs / 60000);



            //var startTime = new Date(parseInt(segment.startTimeNanos, 10) / 1000000);
            //var endTime = new Date(parseInt(segment.endTimeNanos, 10) / 1000000);
            //var duration = (endTime - startTime) / 1000 / 60;

            var sleepStageType = segment.value[0].intVal;
            var sleepStage;
      
            switch(sleepStageType) {
              case 1:
                sleepStage = "Awake (during sleep cycle)";
                break;
              case 2:
                sleepStage = "Sleep";
                break;
              case 3:
                sleepStage = "Out-of-bed";
                break;
              case 4:
                sleepStage = "Light sleep";
                break;
              case 5:
                sleepStage = "Deep sleep";
                break;
              case 6:
                sleepStage = "REM";
                break;
              default:
                sleepStage = "Unknown";
                break;
            }



            //var sleepType = segment.value[0].intVal == 1 ? 'Light' : 'Deep';
            var sleepType = sleepStage;
            var sleepDuration = duration.toFixed(1) + ' min';
            sleepData += sleepType + ' sleep: ' + duration.toFixed(1) + ' min\n';
          }
        }
        
        //sheet.appendRow([bucketDate, 
        //                 sleepData]);
        sheet.appendRow([startTime, endTime, sleepType, 
            sleepDuration]);
      }
  }
  
  