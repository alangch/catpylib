import win32com.client

# To use typename function in VB
# input paramater 'obj' is the object which you want to get typp
def TypeName(obj):
    catia=win32com.client.Dispatch('catia.application')

    code='''
    Function GetTypeName(object)
        GetTypeName = TypeName(object)
    End Function
    '''

    typename = catia.SystemService.Evaluate(code,
                                            0,
                                            'GetTypeName',
                                            [obj])
    return typename



# to measure point coordinates
# input parameter 'mes' is the measurable object
def GetPoint(mes):
    catia=win32com.client.Dispatch('catia.application')

    code='''
    Function GetPoint(mes)
        Dim Arr(2)
        mes.GetPoint Arr
        GetPoint = Arr
    End Function
    '''

    Coord = catia.SystemService.Evaluate(code,
                                            0,
                                            'GetPoint',
                                            [mes])
    return Coord



# to ues 'Nothing' object in VB
# no input needed, function will return a 'Nothing' object
def Nothing():
    catia=win32com.client.Dispatch('catia.application')
    a=1

    code='''
    Function Noth(a)        
        set Noth = Nothing        
    End Function
    '''

    N = catia.SystemService.Evaluate(code,
                                        0,
                                        'Noth',
                                     [a])
    return N