function runDemo() {
  showChooser();
}


function updateAccountInfo(account, property, view, viewName) {
  
  //Store viewName in spreadsheet
  setData(2,8,viewName);
  
  runData(view); 
}


function setData(row,col,val) {
  var ss1 = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss1.getSheets()[0];
  var cell1 = sheet1.getRange(row, col);
  cell1.setValue(val);
  
}


function getViews(account, property, propertyName) {
  var profiles = Analytics.Management.Profiles.list(account, property);
  
  //Store property name in spreadsheet
  setData(2,7,propertyName);
  
  var output = [];
  profiles.items.forEach (function(x) {
    var obj = {"text": x.name, "value": x.id};
    output.push(obj);
  });
  return output;
}


function getProfiles(account,accountName) {
  var profiles = Analytics.Management.Webproperties.list(account);
  
  //Store accountName in spreadsheet
  setData(2,6,accountName);
  
  var output = [];
  profiles.items.forEach (function(x) {
    var obj = {"text": x.name, "value": x.id};
    output.push(obj);
  });
  return output;
}


function getAccountsData() {
  var accounts = Analytics.Management.Accounts.list();
  var output = [];
  accounts.items.forEach (function(x) {
    var obj = {"text": x.name, "value": x.id};
    output.push(obj);
  });
  return output;
}


function getAccountsHTML() {
  var t = HtmlService
      .createTemplateFromFile('AccountsDropdown');
  //t.data = getAccountsData();
  return t.evaluate();
}


function showChooser() {
  // Function: Gets relevant GA profile and returns it in firstProfile
  var html = getAccountsHTML();
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Choose Account');
}


function runData(account, property, view) {
  try {
    
    // Define an array to hold all the information used in each API query and where to return the results in the spreadsheet
    var gaQueries = [];

    // Populate the array with data
    gaQueries.push({
      queryName : "Overall Device Category Breakdown",
      resultRowStart : 7,
      resultColumnStart : 1,
      queryParameters :  
      {
        'dimensions': 'ga:deviceCategory',              
        'sort': '-ga:sessions',                         
        'start-index': '1',
        'max-results': '10'                             
      },  
      specialProcessing : "No"
    });
    
        
    gaQueries.push({
      queryName : "Overall Browser Breakdown",
      resultRowStart : 21,
      resultColumnStart : 1,
      queryParameters :  
      {
        'dimensions': 'ga:browser',              
        'sort': '-ga:sessions',                  
        'start-index': '1',
        'max-results': '10'                    
      },  
      specialProcessing : "No"
    });
    
        
    gaQueries.push({
      queryName : "Overall OS Breakdown",
      resultRowStart : 35,
      resultColumnStart : 1,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystem',      
        'sort': '-ga:sessions',                  
        'start-index': '1',
        'max-results': '10'                      
      },  
      specialProcessing : "No"
    });
    
       
    gaQueries.push({
      queryName : "Desktop OS Breakdown",
      resultRowStart : 7,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystem',                        
        'sort': '-ga:sessions',                                    
        'segment': 'dynamic::ga:deviceCategory==desktop',          
        'start-index': '1',
        'max-results': '10'                                        
      },  
      specialProcessing : "No"
    });    
    
       
    gaQueries.push({
      queryName : "Desktop Windows OS Breakdown",
      resultRowStart : 21,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                     
        'sort': '-ga:sessions',                                      
        'segment': 'dynamic::ga:deviceCategory==desktop',             
        'filters': 'ga:operatingSystem==Windows',                     
        'start-index': '1',
        'max-results': '10'                                            
      },  
      specialProcessing : "No"
    });    
    
        
    gaQueries.push({
      queryName : "Desktop Mac OS Breakdown",
      resultRowStart : 35,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',             
        'sort': '-ga:sessions',                               
        'segment': 'dynamic::ga:deviceCategory==desktop',      
        'filters': 'ga:operatingSystem==Macintosh',             
        'start-index': '1',
        'max-results': '10'  
      },  
      specialProcessing : "Yes",
      specialProcessingParameters :
      [        
        // The values returned from GA is in the form "Intel 1N.N" so the OS version number will start at the 6th index
        {'sortedRowName' : 'macOS 10.14: Mojave',                 'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.13'},
        {'sortedRowName' : 'macOS 10.13: High Sierra',            'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.13'},
        {'sortedRowName' : 'macOS 10.12: Sierra (Fuji)',          'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.12'},
        {'sortedRowName' : 'OS X 10.11: El Capitan (Gala)',       'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.11'},
        {'sortedRowName' : 'OS X 10.10: Yosemite (Syrah)',        'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.10'},
        {'sortedRowName' : 'OS X 10.9 Mavericks (Cabernet)',      'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.9'},
        {'sortedRowName' : 'OS X 10.8 Mountain Lion (Zinfandel)', 'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.8'},
        {'sortedRowName' : 'OS X 10.7 Lion (Barolo)',             'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.7'},
        {'sortedRowName' : 'OS X 10.6 Snow Leopard',              'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.6'},
        {'sortedRowName' : 'OS X 10.5 Leopard (Chablis)',         'sortedRowCount' : 0, 'criteriaStartIndex' : 6, 'sortedRowCountCriteria' : '10.5'}
      ]
    }); 

    
    gaQueries.push({
      queryName : "Desktop Browser Breakdown",
      resultRowStart : 49,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                            
        'sort': '-ga:sessions',                                
        'segment': 'dynamic::ga:deviceCategory==desktop',      
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    }); 


    gaQueries.push({
      queryName : "Desktop Windows Browser Breakdown",
      resultRowStart : 63,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                            
        'sort': '-ga:sessions',                                
        'segment': 'dynamic::ga:deviceCategory==desktop',
        'filters': 'ga:operatingSystem==Windows',
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    });    


    gaQueries.push({
      queryName : "Desktop Mac Browser Breakdown",
      resultRowStart : 77,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                            
        'sort': '-ga:sessions',                                
        'segment': 'dynamic::ga:deviceCategory==desktop',
        'filters': 'ga:operatingSystem==Macintosh',
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    });

    
    gaQueries.push({
      queryName : "Desktop Windows Internet Explorer Breakdown",
      resultRowStart : 91,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:browserVersion',                              
        'sort': '-ga:sessions',                                       
        'segment': 'dynamic::ga:deviceCategory==desktop',              
        'filters': 'ga:browser==Internet Explorer',                   
        'start-index': '1',
        'max-results': '10',
      },  
      specialProcessing : "No"
    }); 
    
        
    gaQueries.push({
      queryName : "Desktop Resolution Breakdown",
      resultRowStart : 105,
      resultColumnStart : 7,
      queryParameters :  
      {
        'dimensions': 'ga:screenResolution',                   
        'sort': '-ga:sessions',                                
        'segment': 'dynamic::ga:deviceCategory==desktop',      
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    }); 
   
        
    gaQueries.push({
      queryName : "Mobile Device Breakdown",
      resultRowStart : 7,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:mobileDeviceInfo',                 
        'sort': '-ga:sessions',                               
        'segment': 'dynamic::ga:deviceCategory==mobile',     
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    });     
    
        
    gaQueries.push({
      queryName : "iPhone Device Breakdown",
      resultRowStart : 21,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:screenResolution',                     
        'sort': '-ga:sessions',                                   
        'segment': 'dynamic::ga:deviceCategory==mobile',         
        'filters': 'ga:mobileDeviceBranding==Apple',             
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "Yes",
      specialProcessingParameters :
      [        
        // Match reported resolutions to known resolutions for models
        // https://www.paintcodeapp.com/news/ultimate-guide-to-iphone-resolutions
        // https://stackoverflow.com/questions/25755443/iphone-6-plus-resolution-confusion-xcode-or-apples-website-for-development
        {'sortedRowName' : 'iPhone(320x480): 2G, 3G, 3GS, 4, 4s',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '320x480'},
        {'sortedRowName' : 'iPhone(320x568): 5, 5s, 5c, SE',                   'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '320x568'},
        {'sortedRowName' : 'iPhone(375x667): 6, 6s, 7, 8',                     'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '375x667'},
        {'sortedRowName' : 'iPhone(414x736): 6 Plus, 6s Plus, 7 Plus, 8 Plus', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '414x736'},
        {'sortedRowName' : 'iPhone(375x812): X, Xs, 11 Pro',                   'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '375x812'},
        {'sortedRowName' : 'iPhone(414x896): Xs Max, XR, 11, 11 Pro Max',      'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '414x896'}
      ]
    }); 


    gaQueries.push({
      queryName : "Mobile Android Device Breakdown",
      resultRowStart : 35,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:mobileDeviceInfo',               
        'sort': '-ga:sessions',                            
        'segment': 'dynamic::ga:deviceCategory==mobile',   
        'filters': 'ga:operatingSystem==Android',          
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    });    

        
    gaQueries.push({
      queryName : "Mobile OS Breakdown",
      resultRowStart : 49,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystem',                      
        'sort': '-ga:sessions',                                  
        'segment': 'dynamic::ga:deviceCategory==mobile',          
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"
    }); 


    gaQueries.push({
      queryName : "Mobile iOS OS Breakdown",
      resultRowStart : 63,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                      
        'sort': '-ga:sessions',                                         
        'segment': 'dynamic::ga:deviceCategory==mobile',      
        'filters': 'ga:operatingSystem==iOS',                          
        'start-index': '1',
        'max-results': '10'   
      },  
      specialProcessing : "No"
    });     
    
       
    gaQueries.push({
      queryName : "Mobile iOS OS Major Version Breakdown",
      resultRowStart : 77,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                       
        'sort': '-ga:sessions',                                         
        'segment': 'dynamic::ga:deviceCategory==mobile',              
        'filters': 'ga:operatingSystem==iOS',                            
        'start-index': '1',
        'max-results': '250'  
      },  
      specialProcessing : "Yes",
      specialProcessingParameters :
      [        
        // The values returned from GA is in the form "NN.N.N"
        {'sortedRowName' : 'iOS 13', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '13'},
        {'sortedRowName' : 'iOS 12', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '12'},
        {'sortedRowName' : 'iOS 11', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '11'},
        {'sortedRowName' : 'iOS 10', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '10'},
        {'sortedRowName' : 'iOS 9',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '9'},
        {'sortedRowName' : 'iOS 8',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '8'},
        {'sortedRowName' : 'iOS 7',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '7'},
        {'sortedRowName' : 'iOS 6',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '6'},
        {'sortedRowName' : 'iOS 5',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '5'},
        {'sortedRowName' : 'iOS 4',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '4'}

      ]      
    });  
    
    
    gaQueries.push({
      queryName : "Mobile Android OS Breakdown",
      resultRowStart : 91,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                
        'sort': '-ga:sessions',                                     
        'segment': 'dynamic::ga:deviceCategory==mobile',            
        'filters': 'ga:operatingSystem==Android',                     
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });     

        
    gaQueries.push({
      queryName : "Mobile Android OS Major Version Breakdown",
      resultRowStart : 105,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                  
        'sort': '-ga:sessions',                                       
        'segment': 'dynamic::ga:deviceCategory==mobile',           
        'filters': 'ga:operatingSystem==Android',                    
        'start-index': '1',
        'max-results': '250' 
      },  
      specialProcessing : "Yes",
      specialProcessingParameters :
      [        
        // The values returned from GA is in the form "NN.N.N"
        {'sortedRowName' : 'Android 10 (10.0+)',                 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '10'},
        {'sortedRowName' : 'Pie (9.0+)',                         'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '9'},
        {'sortedRowName' : 'Oreo (8.0+)',                        'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '8'},
        {'sortedRowName' : 'Nougat (7.0 to 7.1.2)',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '7'},
        {'sortedRowName' : 'Marshmallow (6.0 to 6.1)',           'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '6'},
        {'sortedRowName' : 'Lollipop (5.0 to 5.1)',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '5'},
        {'sortedRowName' : 'KitKat (4.4 to 4.4.4)',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '4.4'},
        {'sortedRowName' : 'Jelly Bean (4.1 to 4.3.1)',          'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : ['4.3','4.2','4.1']},
        {'sortedRowName' : 'Ice Cream Sandwich (4.0 to 4.0.4)',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '4.0'},
        {'sortedRowName' : 'Honeycomb (3.0 to 3.2.6)',           'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : ['3.0','3.1','3.2']},
        //{'sortedRowName' : 'Gingerbread (2.3 to 2.3.7)',         'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '2.3'}
        //{'sortedRowName' : 'Froyo (2.2 to 2.2.3)',               'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '2.2'}
      ]      
    });     


    gaQueries.push({
      queryName : "Mobile Browser Breakdown",
      resultRowStart : 119,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                            
        'sort': '-ga:sessions',                                
        'segment': 'dynamic::ga:deviceCategory==mobile',       
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });    
    
        
    gaQueries.push({
      queryName : "Mobile iOS Browser Breakdown",
      resultRowStart : 133,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                       
        'sort': '-ga:sessions',                            
        'segment': 'dynamic::ga:deviceCategory==mobile',  
        'filters': 'ga:operatingSystem==iOS',   
        'start-index': '1',
        'max-results': '10'  
      },  
      specialProcessing : "No"     
    });  


    gaQueries.push({
      queryName : "Mobile Android Browser Breakdown",
      resultRowStart : 147,
      resultColumnStart : 13,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                    
        'sort': '-ga:sessions',                                    
        'segment': 'dynamic::ga:deviceCategory==mobile',            
        'filters': 'ga:operatingSystem==Android',                       
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });    
    
    
    gaQueries.push({
      queryName : "Tablet Device Breakdown",
      resultRowStart : 7,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:mobileDeviceInfo',       
        'sort': '-ga:sessions',                           
        'segment': 'dynamic::ga:deviceCategory==tablet', 
        'start-index': '1',
        'max-results': '10'
      },  
      specialProcessing : "No"     
    }); 


    gaQueries.push({
      queryName : "iPad Device Breakdown",
      resultRowStart : 21,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:screenResolution',             
        'sort': '-ga:sessions',                           
        'segment': 'dynamic::ga:deviceCategory==tablet',  
        'filters': 'ga:mobileDeviceBranding==Apple',      
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    }); 

        
    gaQueries.push({
      queryName : "Tablet Android Device Breakdown",
      resultRowStart : 35,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:mobileDeviceInfo',        
        'sort': '-ga:sessions',                            
        'segment': 'dynamic::ga:deviceCategory==tablet',   
        'filters': 'ga:operatingSystem==Android',          
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    }); 


    gaQueries.push({
      queryName : "Tablet OS Breakdown",
      resultRowStart : 49,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystem',               
        'sort': '-ga:sessions',                          
        'segment': 'dynamic::ga:deviceCategory==tablet',   
        'start-index': '1',
        'max-results': '10'
      },  
      specialProcessing : "No"     
    });


    gaQueries.push({
      queryName : "Tablet iOS OS Breakdown",
      resultRowStart : 63,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                  
        'sort': '-ga:sessions',                                    
        'segment': 'dynamic::ga:deviceCategory==tablet',            
        'filters': 'ga:operatingSystem==iOS',                         
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });


    gaQueries.push({
      queryName : "Tablet iOS OS Major Version Breakdown",
      resultRowStart : 77,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                
        'sort': '-ga:sessions',                                      
        'segment': 'dynamic::ga:deviceCategory==tablet',              
        'filters': 'ga:operatingSystem==iOS',                      
        'start-index': '1',
        'max-results': '250' 
      },  
      specialProcessing : "Yes",
      specialProcessingParameters :
      [        
        // The values returned from GA is in the form "NN.N.N"
        {'sortedRowName' : 'iOS 13', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '13'},
        {'sortedRowName' : 'iOS 12', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '12'},
        {'sortedRowName' : 'iOS 11', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '11'},
        {'sortedRowName' : 'iOS 10', 'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '10'},
        {'sortedRowName' : 'iOS 9',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '9'},
        {'sortedRowName' : 'iOS 8',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '8'},
        {'sortedRowName' : 'iOS 7',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '7'},
        {'sortedRowName' : 'iOS 6',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '6'},
        {'sortedRowName' : 'iOS 5',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '5'},
        {'sortedRowName' : 'iOS 4',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '4'}
      ]      
    }); 


    gaQueries.push({
      queryName : "Tablet Android OS Breakdown",
      resultRowStart : 91,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',                  
        'sort': '-ga:sessions',                                      
        'segment': 'dynamic::ga:deviceCategory==tablet',            
        'filters': 'ga:operatingSystem==Android',                     
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });


    gaQueries.push({
      queryName : "Tablet Android OS Major Version Breakdown",
      resultRowStart : 105,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:operatingSystemVersion',               
        'sort': '-ga:sessions',                                    
        'segment': 'dynamic::ga:deviceCategory==tablet',            
        'filters': 'ga:operatingSystem==Android',                  
        'start-index': '1',
        'max-results': '250'
      },  
      specialProcessing : "Yes",
      specialProcessingParameters :
      [        
        // The values returned from GA is in the form "NN.N.N" 
        {'sortedRowName' : 'Oreo (8.0+)',                        'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '8'},
        {'sortedRowName' : 'Nougat (7.0 to 7.1.2)',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '7'},
        {'sortedRowName' : 'Marshmallow (6.0 to 6.1)',           'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '6'},
        {'sortedRowName' : 'Lollipop (5.0 to 5.1)',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '5'},
        {'sortedRowName' : 'KitKat (4.4 to 4.4.4)',              'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '4.4'},
        {'sortedRowName' : 'Jelly Bean (4.1 to 4.3.1)',          'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : ['4.3','4.2','4.1']},
        {'sortedRowName' : 'Ice Cream Sandwich (4.0 to 4.0.4)',  'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '4.0'},
        {'sortedRowName' : 'Honeycomb (3.0 to 3.2.6)',           'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : ['3.0','3.1','3.2']},
        {'sortedRowName' : 'Gingerbread (2.3 to 2.3.7)',         'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '2.3'},
        {'sortedRowName' : 'Froyo (2.2 to 2.2.3)',               'sortedRowCount' : 0, 'criteriaStartIndex' : 0, 'sortedRowCountCriteria' : '2.2'}
      ]      
    }); 

    
    gaQueries.push({
      queryName : "Tablet Browser Breakdown",
      resultRowStart : 119,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                    
        'sort': '-ga:sessions',                                     
        'segment': 'dynamic::ga:deviceCategory==tablet',         
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });
    
    
    gaQueries.push({
      queryName : "Tablet iOS Browser Breakdown",
      resultRowStart : 133,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                    
        'sort': '-ga:sessions',                                   
        'segment': 'dynamic::ga:deviceCategory==tablet',         
        'filters': 'ga:operatingSystem==iOS',                     
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });


    gaQueries.push({
      queryName : "Tablet Android Browser Breakdown",
      resultRowStart : 147,
      resultColumnStart : 19,
      queryParameters :  
      {
        'dimensions': 'ga:browser',                    
        'sort': '-ga:sessions',                                    
        'segment': 'dynamic::ga:deviceCategory==tablet',            
        'filters': 'ga:operatingSystem==Android',                       
        'start-index': '1',
        'max-results': '10' 
      },  
      specialProcessing : "No"     
    });

    
    // Sets ss1 as active spreadsheet
    var ss1 = SpreadsheetApp.getActiveSpreadsheet();  
    // Sets sheet1 as first sheet of spreadsheet
    var sheet1 = ss1.getSheets()[0];
    
    // Writes date and time into cell J1 as soon as script is started
    var cell1 = sheet1.getRange(1, 10);
    cell1.setValue(getDateAndTime());
    // Also write it into cell J2 (to stop calc giving negative answer from previous use before script has finished)
    cell1 = sheet1.getRange(2, 10);
    cell1.setValue(getDateAndTime());
    
    // Gets GA profile stored into firstProfile 
    var firstProfile = account;
    
    // Set progress as zero
    cell1 = sheet1.getRange(1, 15);
    cell1.setValue(0/gaQueries.length);
           
    // Set up variable to hold arguments for GA query (do I need to initialise here?)
    var optArgs = {
      'dimensions': 'ga:mobileDeviceInfo',               // Comma separated list of dimensions.
      'sort': '-ga:sessions',                            // Sort by sessions descending, then keyword.
      'segment': 'dynamic::ga:deviceCategory==tablet',   // Process only tablet traffic.
      'filters': 'ga:operatingSystem==Android',          // Display only google traffic.
      'start-index': '1',
      'max-results': '10'
    };


    
    // Main query loop
    for (var i = 0; i < gaQueries.length ; ++i)
    {
      
      // Read query name from query object and write into results table
      cell1 = sheet1.getRange(gaQueries[i].resultRowStart - 1, gaQueries[i].resultColumnStart);
      cell1.setValue(gaQueries[i].queryName);
      
      // Write query name in processing box
      cell1 = sheet1.getRange(2, 13);
      cell1.setValue(gaQueries[i].queryName);    
      
      // Update progress % cell on spreadsheet
      cell1 = sheet1.getRange(1, 15);
      cell1.setValue((i+1)/gaQueries.length);    
            
      // If writing additional parameters into the query do it here 
      // Set filterString as the additional filter you want 
      // e.g. filtering for suppliers 'ga:pagePath=@/supplier/'
      // or filtering for venues 'ga:pagePath=@/venuesadmin/;ga:pagePath=@rid'
      
      //Read in if there are any additional filters set in the spreadsheet (cell Q17)
      var filter = "";
      cell1 = sheet1.getRange(2,17);
      filter = cell1.getValue();
      
      // Variable for the filter string
      var filterString = "";
      
      //Set the filter string accordingly
      if (filter === "Venues") {filterString = "ga:pagePath=@/venuesadmin/;ga:pagePath=@rid";} //venues
      if (filter === "Suppliers") {filterString = "ga:pagePath=@/supplier/";} //suppliers 
      if (filter === "PlanningTools" ) {filterString = "ga:pagePath=@/planner/";} //planning tools

      //var filterString = "" ;//Nothing
      //var filterString = "ga:pagePath=@/venuesadmin/;ga:pagePath=@rid"; //venues
      //var filterString = "ga:pagePath=@/supplier/"; //suppliers
      
      // Dont want to run this code is no additional filter is set
      if (filterString !== "")
      {
        // Is current filter setfor the query?
        if (typeof gaQueries[i].queryParameters.filters === "undefined")
        {
          // If not set the filter
          gaQueries[i].queryParameters.filters = filterString;
          //Browser.msgBox("Newly set " + gaQueries[i].queryParameters.filters);
        }
        else
        {
          // If so append the filter
          gaQueries[i].queryParameters.filters = gaQueries[i].queryParameters.filters + ";" + filterString;
          //Browser.msgBox("Newly set " + gaQueries[i].queryParameters.filters);
        }
      }
      
      // Read the GA results into results
      results = getReportDataForProfile(firstProfile, gaQueries[i].queryParameters);
      
      // Do any special processing here
      if (gaQueries[i].specialProcessing == "Yes")
          {
            //Do the processing
            doSpecialProcessing(results, i, gaQueries);
          }
      
      // Write the query values to the spreadsheet
      outputToSpreadsheet(results, gaQueries[i].resultRowStart, gaQueries[i].resultColumnStart);
      
      // Write in calculations
      //if ( i===0 /*|| i === 1*/)
      {
        setCellCalculations(i, gaQueries);
      }
      
    }
    
    // Write into processing box in spreadsheet
    cell1 = sheet1.getRange(2, 13);
    cell1.setValue("Complete");
    
    // Write finish time into spreadsheet
    cell1 = sheet1.getRange(2, 10);
    cell1.setValue(getDateAndTime());
    
    //Browser.msgBox("Finished");
    
    
  } catch(error) {
    Browser.msgBox(error.message);
  }
}



function getDate() {
  // Function: Returns Date in DD/MM/YYYY format
  var d = new Date();
  return (d.getDate()) + "/" + (d.getMonth()+1) + "/" + d.getFullYear();
}



function getTime() {
  // Function: Returns Time in hh:mm:ss format
  var d = new Date(),
      offset = -d.getTimezoneOffset()/60,
      h = d.getUTCHours() + offset,
      m = d.getMinutes(),
      s = d.getSeconds();
  return h + ":" + m + ":" + s;
}



function getDateAndTime() {
  // Function: Concatenation of date and time functions
  return getDate() + " " + getTime();
}



function getFirstProfile() {
  // Function: Gets relevant GA profile and returns it in firstProfile
  doGet();
}



function getReportDataForProfile(firstProfile, optArgs) {

  var profileId = firstProfile;
  var tableId = 'ga:' + profileId;
  
  // Set ss as active spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 

  // Set first sheet as the active sheet
  var sheet = ss.getSheets()[0]; 

  // Read in dates to 2D array from spreadsheet
  var dates = sheet.getSheetValues(1, 1, 2, 2) 
  
  // Set start date
  startDate = dates[0][1]; 
  // Convert date into correct string/format (needed for GA api query)  
  startDate = Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd'); 
    
  // Set end date
  endDate = dates[1][1];
  // Convert date into correct string/format (needed for GA api query) 
  endDate = Utilities.formatDate(endDate, 'GMT', 'yyyy-MM-dd'); 
   
  // Make a request to the API.
  var results = Analytics.Data.Ga.get(
      tableId,                    // Table id (format ga:xxxxxx).
      startDate,                  // Start-date (format yyyy-MM-dd).
      endDate,                    // End-date (format yyyy-MM-dd).
      'ga:sessions', // Comma seperated list of metrics.
      optArgs);
  
  // This checks if there are any rows and throws an error about profiles if not (however this logic maybe wrong as sometimes there may be no rows but the profile in use is perfectly fine shit code google!)
  if (results.getRows()) {
    return results;

  } else {
    results.rows = [["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""]];
    return results;
    //throw new Error('No views (profiles) found');
  }
  
}



function outputToSpreadsheet(results, tableRowStart, columnRowStart) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    
  // Print the headers.
  var headerNames = []; //Array of header names
  
  // Fill array with header names from GA query (results)
  for (var i = 0, header; header = results.getColumnHeaders()[i]; ++i) {
    headerNames.push(header.getName());
  }
  
  // Write the header names to the spreadsheet
  sheet.getRange(tableRowStart, columnRowStart, 1, headerNames.length)
       .setValues([headerNames]);
  
  // Blank previous results 
  var blankRows = [["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""]];
  sheet.getRange(tableRowStart + 1, columnRowStart, blankRows.length, headerNames.length)
       .setValues(blankRows);
  
  // Write the rows of data.
  sheet.getRange(tableRowStart + 1, columnRowStart, results.getRows().length, headerNames.length)
       .setValues(results.getRows());

  
  
  // Write the total session in the query
  sheet.getRange(tableRowStart - 1, columnRowStart + 4, 1, 1)
       .setValue(String(results.totalsForAllResults['ga:sessions']));  
  
}



function doSpecialProcessing(results, i, gaQueries) {
  
  // Set up variable for sorted list
  var sortedResults = [];
  
  // Temporary variable for results
  // Could be any length at this point
  var resultsTemp = results.getRows();
  
  //Browser.msgBox("in special processing function");
  
  // Loop to build sorted list
  for (var j = 0 ; j < gaQueries[i].specialProcessingParameters.length ; ++j) 
  {    
    // Populate sorted list from query object
    // [0] row names (string)
    // [1] row count (initially zero)
    // [2] row criteria (might be an array of strings or just a string)
    // [3] criteriaStartIndex (an int)
    sortedResults.push([gaQueries[i].specialProcessingParameters[j].sortedRowName,
                        gaQueries[i].specialProcessingParameters[j].sortedRowCount, 
                        gaQueries[i].specialProcessingParameters[j].sortedRowCountCriteria,
                        gaQueries[i].specialProcessingParameters[j].criteriaStartIndex
                       ]);    
  }
    
  // Loop through results rows (might be 250 of them)
  for (var j = 0 ; j < resultsTemp.length ; ++j)
  {  
    // Loop through the substrings (usually 10 but not always, never more than 10)
    for (var k = 0 ; k < sortedResults.length ; ++k)
    {  
      
      // Need to loop through array if one exists
      // But only add on first hit then stop (as am doing ORs)
      // Check if the row contains the relevant substring(s) (could be an array of them)
      if (resultsTemp[j][0].substring(sortedResults[k][3],sortedResults[k][3] + sortedResults[k][2].length) === sortedResults[k][2]) 
      {
        // If it does add the quantity from the row to the sorted results        
        sortedResults[k][1] += parseInt(resultsTemp[j][1]);
      }  
      
      //Check if the sort criteria is an array
      // If it is (or not a string) need to check the second string (element) onwards (first has already been checked above)
      if (typeof sortedResults[k][2] !== "string")
      {
        // boolean to check if we have hit a match
        var match = false;
        // counter to loop through strings in array
        var l = 1;
        
        // while we havent had a match or hit the end of the array
        while (match === false && l < sortedResults[k][2].length)
        {          
          // Check string in array
          if (resultsTemp[j][0].substring(sortedResults[k][3],sortedResults[k][3] + sortedResults[k][2].length) === sortedResults[k][2][l])
          {
            //Browser.msgBox("checking for " + sortedResults[k][2][l]);
            match = true;
            sortedResults[k][1] += parseInt(resultsTemp[j][1]);
          }
          // Increase index to check next string
          ++l;
        }
        
      }
      
    }  
  }
    
  // Resize and initialise resultsTemp to 10 results
  resultsTemp = [["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""],["",""]];
    
  // Re-order the sort results based on quantity  
  sortedResults.sort(compareSecondColumn);

  function compareSecondColumn(a, b) {
    if (a[1] === b[1]) {
        return 0;
    }
    else {
        return (b[1] < a[1]) ? -1 : 1;
    }
  }
  
  // Write sorted results back into resultsTemp ready for outputting to spreadsheet
  for (var j = 0 ; j < sortedResults.length ; ++j)
  {    
    if (sortedResults[j][1] !== 0) 
    {
      // Only add if non zero
      resultsTemp[j][0] = sortedResults[j][0];
      resultsTemp[j][1] = sortedResults[j][1];
    }
    else
    {
      // if zero blank row out
      resultsTemp[j][0] = "";
      resultsTemp[j][1] = "";
    }    
  }
  
  results.rows = resultsTemp;
}  


// implement this later to write the correct formula in each cacluated cell in each results table
// I.e. want something like this as an example =IFERROR(IF(B42/B46=0," ",B42/B46))
// This should ultimately blank out any cells that are blank due to being zero or blank (i.e. the function above)
// Maybe dont need the IFERROR( bit
// Do I want a space in the " " bit? maybe I just want ""
function setCellCalculations(i, gaQueries) {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  // gaQueries.resultRowStart (7 for first table) calculations start at row 8 (and there are 10 rows)
  // bottom row of table is some sums (this will be gaQueries[i].resultRowStart + 11)
  // column 1 is just the label total (this will be gaQueries[i].resultColumnStart, col A, G, M, S)
  // column 2 is the total of session in the query (this will be gaQueries[i].resultColumnStart + 1, col B, H, N, T)
  // column 3 is the total of the top 10 percentages (this will be gaQueries[i].resultColumnStart + 2, col C, I, O, U)
  // column 4 is the total of the query percentages (this will be gaQueries[i].resultColumnStart + 3, col D, J, P, V)
  // column 5 is the total of the total session percentages (this will be gaQueries[i].resultColumnStart + 4, col E, K, Q, W)
  // gaQueries.resultColumnStart (1 for first table) calculations start at column 3 (%top 10) 4 (%query) 5 (%all trafic) 
  // column 3 (%top 10) - divisor is table total sessions i.e. the first total calculated
  // column 4 (%query) - divisor is query total session i.e. top right of table
  // column 5 (%all trafic) - divisor is query total session from overall traffic query i.e. always cell B18
  
  // Write totals for table
  
  // Write in total label
  // Have a var to store contrsucted string to past to cell
  var cellValue = "Total";
  
  // Write value to cell
  sheet.getRange(gaQueries[i].resultRowStart + 11, gaQueries[i].resultColumnStart, 1, 1)
       .setValue(cellValue);
  
  // Have value to use in string contrcution
  var sumStart = gaQueries[i].resultRowStart + 1;
  var sumFinish =  gaQueries[i].resultRowStart + 10;   
  
  // Debug
  //Browser.msgBox("columnToLetter(gaQueries[i].resultColumnStart + 1) = " + columnToLetter(gaQueries[i].resultColumnStart + 1) + " columnToLetter(gaQueries[i].resultColumnStart + 2) = " + columnToLetter(gaQueries[i].resultColumnStart + 2));
  
  // Write in total query sessions, construct srting and write into cell
  cellValue = "=SUM(" + columnToLetter(gaQueries[i].resultColumnStart + 1) + sumStart.toString() + ":" + columnToLetter(gaQueries[i].resultColumnStart + 1) + sumFinish.toString() +")";
  sheet.getRange(gaQueries[i].resultRowStart + 11, gaQueries[i].resultColumnStart + 1, 1, 1)
       .setValue(cellValue);
  
  // Write in total of the top 10 percentages
  cellValue = "=SUM(" + columnToLetter(gaQueries[i].resultColumnStart + 2) + sumStart.toString() + ":" + columnToLetter(gaQueries[i].resultColumnStart + 2) + sumFinish.toString() +")";
  sheet.getRange(gaQueries[i].resultRowStart + 11, gaQueries[i].resultColumnStart + 2, 1, 1)
       .setValue(cellValue);
  
  // Write in total of the query percentages
  cellValue = "=SUM(" + columnToLetter(gaQueries[i].resultColumnStart + 3) + sumStart.toString() + ":" + columnToLetter(gaQueries[i].resultColumnStart + 3) + sumFinish.toString() +")";
  sheet.getRange(gaQueries[i].resultRowStart + 11, gaQueries[i].resultColumnStart + 3, 1, 1)
       .setValue(cellValue);
  
  // Write in the total session percentages
  cellValue = "=SUM(" + columnToLetter(gaQueries[i].resultColumnStart + 4) + sumStart.toString() + ":" + columnToLetter(gaQueries[i].resultColumnStart + 4) + sumFinish.toString() +")";
  sheet.getRange(gaQueries[i].resultRowStart + 11, gaQueries[i].resultColumnStart + 4, 1, 1)
       .setValue(cellValue);
  
  // Write in cell calulations
  // Loop for 10 rows of table
  for (var j = 0 ; j < 10 ; ++j)
  {
    // Write in %top 10 column calcs
     
    // =IFERROR(IF(B42/B46=0," ",B42/B46))
    var rowWriteA = gaQueries[i].resultRowStart + j + 1;
    var rowWriteB = gaQueries[i].resultRowStart + 11;
    var division = columnToLetter(gaQueries[i].resultColumnStart + 1) + rowWriteA + "/" + columnToLetter(gaQueries[i].resultColumnStart + 1) + rowWriteB;
    cellValue = "=IFERROR(IF(" + division + "=0,\" \"," + division + "))";
    
    sheet.getRange(gaQueries[i].resultRowStart + j + 1, gaQueries[i].resultColumnStart + 2, 1, 1)
         .setValue(cellValue);
    
    //Write in (%query) column calcs
    rowWriteB = gaQueries[i].resultRowStart - 1;
    division = columnToLetter(gaQueries[i].resultColumnStart + 1) + rowWriteA + "/" + columnToLetter(gaQueries[i].resultColumnStart + 4) + rowWriteB;
    cellValue = "=IFERROR(IF(" + division + "=0,\" \"," + division + "))";
    
    sheet.getRange(gaQueries[i].resultRowStart + j + 1, gaQueries[i].resultColumnStart + 3, 1, 1)
         .setValue(cellValue);
    
    //Write in (%query) column calcs
    division = columnToLetter(gaQueries[i].resultColumnStart + 1) + rowWriteA + "/B18";
    cellValue = "=IFERROR(IF(" + division + "=0,\" \"," + division + "))";
    
    sheet.getRange(gaQueries[i].resultRowStart + j + 1, gaQueries[i].resultColumnStart + 4, 1, 1)
         .setValue(cellValue);
  }
  
}


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
