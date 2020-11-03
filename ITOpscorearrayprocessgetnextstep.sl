namespace: com.ust.clones

operation:
  name: ITOpscorearrayprocessgetnextstep

  inputs:
    - arrayindex:
        default: '0'
        required: false
    - totallength:
        default: '0'
        required: false
  python_action:
    script: |
      intArrayIndex = int(arrayindex)
      intTotalLength = int(totallength)
      nextcounter = 0
      if ((intArrayIndex+1) == intTotalLength):
        return_code = 0
      else:
        nextcounter = intArrayIndex + 1
        return_code = 1
  outputs:
    - return_code: ${ str(return_code) }
    - next_counter: ${str(nextcounter)}
    - totalsize: ${str(totallength)}
  results:
    - SUCCESS: ${return_code == 1}
    - FAILURE