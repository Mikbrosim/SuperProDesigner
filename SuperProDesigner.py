from IPython import embed
import comtypes.client
from comtypes.automation import VARIANT, byref

class SuperProDesigner:
    def __init__(self,tlb_path = r"C:\Program Files (x86)\Intelligen\SuperPro Designer\v10\Designer.tlb"):
        self.Designer = comtypes.client.GetModule(tlb_path)
        self.app = comtypes.client.CreateObject(self.Designer.Application)

    #Application Related Methods:
    """
    These methods are used for performing general application tasks such as activating the designer, application, opening and closing files, etc.
    """
    def ShowApp(self):
        # ShowApp( ) This subroutine is used to activate the Pro-Designer application and display it in its current size position.
        return self.app.ShowApp()
    def CloseApp(self):
        # CloseApp( ) This subroutine is used to close the Pro-Designer application. If there are Pro-Designer case files still open it will close all the documents without saving them.
        return self.app.CloseApp()
    def OpenDoc(self,fileName:str):
        # OpenDoc(fileName As String) This function is used to open the Pro-Designer file with name fileName, makes this file the active Document object, and returns a reference to the caller.
        self.doc = self.app.OpenDoc(fileName)
        return self.doc
    def SetActiveDoc(self,fileName:str):
        raise NotImplementedError()
        # SetActiveDoc(fileName As String) This function is used to activate the Pro-Designer file with name fileName and also returns a reference to this file as a Document object.
        return self.app.SetActiveDoc(fileName)
    def CloseAllDocs(self,bSaveIfNeeded:bool):
        raise NotImplementedError()
        # CloseAllDocs(bSaveIfNeeded As Boolean) This subroutine is used to close all open Pro-Designer file (Document objects) Use bSaveIfNeeded = True for saving the Designer case files and bSaveIfNeeded = False for just closing the documents.
        return self.app.CloseAllDocs(bSaveIfNeeded)

    #Document Related Methods:
    """
    These methods are used for performing generic document tasks on specific Pro-Designer case files.
    """
    def CloseDoc(self,bSaveIfNeeded:bool):
        #CloseDoc(bSaveIfNeeded As Boolean) This subroutine is used to close the active Pro-Designer file (Document object). Use bSaveIfNeeded = True for saving the Designer case file and bSaveIfNeeded = False for just closing the document.
        return self.doc.CloseDoc(bSaveIfNeeded)
    def GetDocName(self):
        raise NotImplementedError()
        #GetDocName(fileName As String, nMaxChar As Long) This function is used to return the name of the active Pro-Designer file (Document object). The function returns a Boolean which is True if it was successful in obtaining the file name and False if it was not. The filename argument is an output argument and returns the name of the Pro-Designer file. The nMaxChar is an input argument and specifies the number of characters that the file name will contain
        doc_name = self.doc.GetDocName()
        return doc_name
    def SaveDoc(self):
        #SaveDoc() This subroutine is used to save the the active Document object.
        return self.doc.SaveDoc()

    #Simulation Related Methods
    """
    These methods are used for simulation tasks. They are all functions that return a Boolean value, which is True if the task was successful and False if the task failed. They include:
    """
    def DoMEBalances(self):
        #DoMEBalances(val) This function is equivalent to clicking on the Solve button or to selecting Tasks / Do M&E Balances from the Pro-Designer application main menu. The value of variable (val) is currently of no importance. 
        return self.doc.DoMEBalances(byref(VARIANT()))

    def DoEconomicCalculations(self):
        #DoEconomicCalculations( ) This function is equivalent to selecting Tasks / Perform Economic Calculations from the Pro-Designer application main menu. 
        return self.doc.DoEconomicCalculations()
    def ScaleUpThroughput(self):
        raise NotImplementedError()
        #ScaleUpThroughput(VarID As VarID, val) This function is used for scaling the process throughput (It is equivalent to selecting Tasks / Adjust Process Throughput from the Pro-Designer application main menu and selecting the Based on Scale Up / Down Factor option). Use VarID = scaleUpFactor_VID and the value of the scale up factor for val (val is a Variant, it’s type should be double and its value should be greater than zero). 
        return self.doc.ScaleUpThroughput()


    # Functions for Process (Flowsheet) Variables
    # Functions for Section Variables
    # Functions for Procedure Variables
    # Functions for Equipment Variables
    # Functions for Operation Variables
    
    # Functions for Stream Variables
    """
    The following functions can be used for setting or retrieving variables that refer to a specific stream (input /output /intermediate) that is included in the process file:
    """
    def GetStreamVarVal(self,streamName:str, VarID:int, compLocalName:str):
        # GetStreamVarVal(streamName As String, VarID As VarID, val, compLocalName As String) can be used to retrieve the value of input/output variables related to the specific stream
        # VarID - Example "Designer.massFlow_VID"
        out_var = VARIANT()
        if not self.doc.GetStreamVarVal(streamName, VarID, byref(out_var), compLocalName): return False
        assert isinstance(out_var.value,float),type(out_var)
        return out_var.value

    def SetStreamVarVal(self, streamName:str, VarID:int, val:float|int, compLocalName:str) -> bool:
        # SetStreamVarVal(streamName As String, VarID As VarID, val, compLocalName As String) can be used for setting input variables related to the specific stream
        # VarID - Example "Designer.massFlow_VID"
        assert isinstance(val,(float,int))
        return self.doc.SetStreamVarVal(streamName, VarID, float(val), compLocalName)

    def AddIngredientToInputStream(self):
        raise NotImplementedError()
        # AddIngredientToInputStream(streamName As String, ingredientName As String, VarID As VarID, val) can be used to add pure components and/or stock mixtures as well as the ingredient’s mass/mole flow or mass fraction to an input stream. The variable IDs that can be used with this function are: componentMassFlow_VID, componentMoleFlow_VID or compMassFrac_VID.
        return self.doc.AddIngredientToInputStream()

    def RemoveIngredientFromInputStream(self,streamName:str,ingredientName:str):
        raise NotImplementedError()
        # RemoveIngredientFromInputStream(streamName As String, ingredientName As String) can be used to remove an ingredient from an input stream.
        return self.doc.RemoveIngredientFromInputStream(streamName,ingredientName)

    def IsInputStreamCompositionValid(self,streamName:str):
        # IsInputStreamCompositionValid(streamName As String) can be used to validate the composition of an input stream. The sum of all ingredient mass fractions should add-up to 1.0.
        return self.doc.IsInputStreamCompositionValid(streamName)


    # Functions for Ingredient Variables
    # Functions for Heat Transfer Agent Variables
    # Functions for Power Variables
    # Functions for Report Option Variables
    # Functions for Excel Data Link Variables
    # Functions for Excel Table Variables

if __name__ == "__main__":
    spd = SuperProDesigner()
    superProDoc = spd.OpenDoc(r"C:\Users\mikbr\Desktop\SPD\COM\v1.spf")
    print(superProDoc)

    print(spd.GetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,""))
    print(spd.SetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,10,""))
    print(spd.GetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,""))
    print(spd.SetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,20,""))
    print(spd.GetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,""))
    print(spd.DoMEBalances())
    embed()
    spd.CloseDoc(False)
    spd.CloseApp()