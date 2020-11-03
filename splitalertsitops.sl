namespace: com.ust.clones

operation:
  name: splitalertsitops

  inputs:
    - text:
        default: ''
        required: false

  python_action:
    script: |
      import json

      def sendjson(alert, index):
       try:
         return alert
       except:
         print("Exception")	  

      jsonobj = json.loads(text)
      index = jsonobj['index']
      inputalerts = jsonobj ['inputJSON']
      for alert in inputalerts:
        alert = sendjson(alert, index)     	  
      

      return_code = 0
  outputs:
    - return_code: ${ str(return_code) }  
    - return_val: ${index}
  results:
    - SUCCESS: ${return_code == 0}
    - FAILURE