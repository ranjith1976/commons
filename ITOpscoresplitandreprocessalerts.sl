namespace: com.ust.clones

operation:
  name: ITOpscoresplitandreprocessalerts

  inputs:
    - organizationId:
        default: ''
        required: true
    - projectId:
        default: ''
        required: true
    - text:
        default: ''
        required: false
    - arrayindex:
        default: '0'
        required: false
  python_action:
    script: |
      import json

      def sendJson(alert, index):
       try:
         print alert
       except:
         print("Exception")	  

      def buildJson(alert,index):
       alertObj = {}
       alertObj['organizationId'] = organizationId
       alertObj['automationStoryName'] = 'PwfITOpsRealtime'
       alertObj['projectId'] = projectId
       alertObj['requestJson'] = {}
       alertObj['requestJson']['index'] = index
       alertObj['requestJson']['inputJSON'] = []
       alertObj['requestJson']['inputJSON'].append(alert)
       alertstr = json.dumps(alertObj)
       return alertstr


      jsonobj = json.loads(text)
      index = jsonobj['index']
      inputalerts = jsonobj ['inputJSON']
      arraysize = len(inputalerts)
      alertobj = None
      alert = inputalerts[int(arrayindex)]
      alertobj = buildJson(alert, index)
      return_code = 0
  outputs:
    - return_code: ${ str(return_code) }  
    - return_val: ${index}
    - alert: ${alertobj}
    - total: ${str(arraysize)}
    - currentindex: ${str(arrayindex)}
  results:
    - SUCCESS: ${return_code == 0}
    - FAILURE