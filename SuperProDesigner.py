from IPython import embed
import comtypes.client
from comtypes.automation import VARIANT, byref

class SuperProDesigner:
    def __init__(self,tlb_path = r"C:\Program Files (x86)\Intelligen\SuperPro Designer\v10\Designer.tlb"):
        Designer = self.Designer = comtypes.client.GetModule(tlb_path)
        self.app = comtypes.client.CreateObject(self.Designer.Application)
        
        # Enumerators
        self.enum_vars:dict[str,dict[str,tuple[int,int]]] = {
            "flowsheet_CID": {
                "unitProc_LID": (Designer.flowsheet_CID, Designer.unitProc_LID),  # Unit Procedure
                "equipment_LID": (Designer.flowsheet_CID, Designer.equipment_LID),  # Equipment
                "stream_LID": (Designer.flowsheet_CID, Designer.stream_LID),  # Streams
                "inStream_LID": (Designer.flowsheet_CID, Designer.inStream_LID),  # Input Streams
                "outStream_LID": (Designer.flowsheet_CID, Designer.outStream_LID),  # Output Streams
                "pureComp_LID": (Designer.flowsheet_CID, Designer.pureComp_LID),  # Pure Components
                "stockMix_LID": (Designer.flowsheet_CID, Designer.stockMix_LID),  # Stock Mixtures
                "mainBranchSection_LID": (Designer.flowsheet_CID, Designer.mainBranchSection_LID),  # Main Branch Sections
                "branch_LID": (Designer.flowsheet_CID, Designer.branch_LID),  # Branches
                "labor_LID": (Designer.flowsheet_CID, Designer.labor_LID),  # Labors
                "hxAgent_LID": (Designer.flowsheet_CID, Designer.hxAgent_LID),  # Heat Transfer Agent
                "power_LID": (Designer.flowsheet_CID, Designer.power_LID),  # Power
                "consumable_LID": (Designer.flowsheet_CID, Designer.consumable_LID),  # Consumables
                "storageUnit_LID": (Designer.flowsheet_CID, Designer.storageUnit_LID)  # Storage units
            },
            "branch_CID": {
                "section_LID": (Designer.branch_CID, Designer.section_LID)  # Sections
            },
            "mainBranchSection_CID": {
                "unitProc_LID": (Designer.mainBranchSection_CID, Designer.unitProc_LID)  # Unit Procedures
            },
            "equipment_CID": {
                "unitProc_LID": (Designer.equipment_CID, Designer.unitProc_LID),  # Unit Procedures
                "variableId_LID": (Designer.equipment_CID, Designer.variableId_LID),  # Equip. Variable Ids
                "staggeredEquip_LID": (Designer.equipment_CID, Designer.staggeredEquip_LID)  # Staggered Equipment
            },
            "unitProc_CID": {
                "operation_LID": (Designer.unitProc_CID, Designer.operation_LID),  # Operations
                "inStream_LID": (Designer.unitProc_CID, Designer.inStream_LID),  # Input Streams
                "outStream_LID": (Designer.unitProc_CID, Designer.outStream_LID)  # Output Streams
            },
            "operation_CID": {
                "reaction_LID": (Designer.operation_CID, Designer.reaction_LID),  # Reactions
                "cleanStep_LID": (Designer.operation_CID, Designer.cleanStep_LID)  # CIP Cleaning Steps
            },
            "stream_CID": {
                "pureComp_LID": (Designer.stream_CID, Designer.pureComp_LID),  # Pure Components
                "stockMix_LID": (Designer.stream_CID, Designer.stockMix_LID),  # Stock Mixtures
                "sourceOperation_LID": (Designer.stream_CID, Designer.sourceOperation_LID),  # Source Operation
                "destinationOperation_LID": (Designer.stream_CID, Designer.destinationOperation_LID)  # Destination Operation
            },
            "stockMix_CID": {
                "pureComp_LID": (Designer.stockMix_CID, Designer.pureComp_LID)  # Pure Components
            }
        }

        # Dictionary for variables relevant to input streams
        self.stream_vars:dict[str,int] = {
            "temperature_VID": Designer.temperature_VID,  # Stream Temperature
            "pressure_VID": Designer.pressure_VID,        # Stream Pressure
            "streamPrice_VID": Designer.streamPrice_VID,  # Stream Price
            "comments_VID": Designer.comments_VID,        # User Comments
            "activity_VID": Designer.activity_VID,        # Stream Activity
            "massFlow_VID": Designer.massFlow_VID,        # Stream Mass Flow
            "volFlow_VID": Designer.volFlow_VID,          # Stream Volumetric Flow
            "mw_VID": Designer.mw_VID,                    # Specified Component Molecular Weight
            "componentMassFlow_VID": Designer.componentMassFlow_VID, # Specified Component Mass Flow in Stream
            "compMassFrac_VID": Designer.compMassFrac_VID,           # Specified Component Mass Fraction in Stream
            "componentMoleFlow_VID": Designer.componentMoleFlow_VID, # Specified Component Mole Flow in Stream
            "compMoleFrac_VID": Designer.compMoleFrac_VID,           # Specified Component Mole Fraction in Stream
            "compExtraCellFrac_VID": Designer.compExtraCellFrac_VID, # Specified Component Extra Cellular Fraction in Stream
            "compVaporFrac_VID": Designer.compVaporFrac_VID,         # Specified Component Vapor Fraction in Stream
            "enthalpy_VID": Designer.enthalpy_VID,                  # Enthalpy of Stream
            "specificEnthalpy_VID": Designer.specificEnthalpy_VID,  # Specific Enthalpy of Stream
            "Cp_VID": Designer.Cp_VID,                              # Heat Capacity of Stream
            "isInputStream_VID": Designer.isInputStream_VID,        # Check if a Stream is an Input Stream
            "isOutputStream_VID": Designer.isOutputStream_VID,      # Check if a Stream is an Output Stream
            "isRawMaterial_VID": Designer.isRawMaterial_VID,        # Is it a "Raw Material"?
            "isCleaningAgent_VID": Designer.isCleaningAgent_VID,      # Is it a "Cleaning Agent"?
            "isMainRevenue_VID": Designer.isMainRevenue_VID,        # Is it a "Main Revenue"?
            "isRevenue_VID": Designer.isRevenue_VID,                # Is it a "Revenue"?
            "isWaste_VID": Designer.isWaste_VID,                    # Is it a "Waste"?
            "isSolidWaste_VID": Designer.isSolidWaste_VID,          # Is it a "Solid Waste"?
            "isCredit_VID": Designer.isCredit_VID,                  # Is it a "Credit"?
            "isAqueousWaste_VID": Designer.isAqueousWaste_VID,      # Is it an "Aqueous Waste"?
            "isOrganicWaste_VID": Designer.isOrganicWaste_VID,      # Is it an "Organic Waste"?
            "isEmission_VID": Designer.isEmission_VID,              # Is it an "Emission"?
            "isNone_VID": Designer.isNone_VID,                      # Is it Classified as "None"?
            "classification_VID": Designer.classification_VID,      # Stream Classification
            "wasteTreatCost_VID": Designer.wasteTreatCost_VID,      # Waste Treatment Cost
            "compMassConc_VID": Designer.compMassConc_VID,                # Specified Ingredient Mass Concentration in Stream
            "compMoleConc_VID": Designer.compMoleConc_VID,                # Specified Ingredient Mole Concentration in Stream
            "autoAdjust_VID": Designer.autoAdjust_VID,                    # Is the Stream Flow Auto Adjusted?
            "bEditIngredientFracs_VID": Designer.bEditIngredientFracs_VID,# Do We Edit the Ingredient Fractions?
            "bVolFlowSetByUser_VID": Designer.bVolFlowSetByUser_VID       # Do We Edit the Stream Mass Flow?
        }
        
        self.procedure_vars:dict[str,int] = {
            "numberOfOperations_VID": Designer.numberOfOperations_VID,  # Number of Operations in the Procedure
            "numberOfCycles_VID": Designer.numberOfCycles_VID,          # Number of Cycles in the Procedure
            "startTime_VID": Designer.startTime_VID,                    # Start Time
            "endTime_VID": Designer.endTime_VID,                        # End Time
            "cycleTime_VID": Designer.cycleTime_VID,                    # Cycle Time
            "holdupTime_VID": Designer.holdupTime_VID,                  # Holdup Time
            "totalTimePerBatch_VID": Designer.totalTimePerBatch_VID,    # Total Time per Batch (all cycles)
            "isBatchMode_VID": Designer.isBatchMode_VID,                # Is Batch Mode?
            "equipmentName_VID": Designer.equipmentName_VID,            # Equipment Name
            "sizeUtilization_VID": Designer.sizeUtilization_VID,        # Size Utilization
            "maxFillRatio_VID": Designer.maxFillRatio_VID,              # Maximum Fill Ratio
            "timeUtilization_VID": Designer.timeUtilization_VID,        # Time Utilization
            "description_VID": Designer.description_VID,                # Description
            "comments_VID": Designer.comments_VID                       # Comments
        }
        self.equipment_vars:dict[str,int] = {
            "noUnits_VID": Designer.noUnits_VID,  # Number of Units
            "noHostedProcedures_VID": Designer.noHostedProcedures_VID,  # Number of Procedures hosted by this equipment
            "isDesignMode_VID": Designer.isDesignMode_VID,  # Is Equipment In Design Mode?
            "noStaggeredEquip_VID": Designer.noStaggeredEquip_VID,  # Number of Staggered Equipment Sets
            "equipPC_VID": Designer.equipPC_VID,  # Purchase Cost
            "equipPCEstimateOption_VID": Designer.equipPCEstimateOption_VID,  # Purchase Cost Estimation Option
            "equipStandByNoUnits_VID": Designer.equipStandByNoUnits_VID,  # Number of Standby Units
            "equipPCDeprecPortion_VID": Designer.equipPCDeprecPortion_VID,  # PC Portion Already Depreciated
            "equipConstrMaterial_VID": Designer.equipConstrMaterial_VID,  # Construction Material
            "equipConstrMaterialF_VID": Designer.equipConstrMaterialF_VID,  # Construction Material Factor
            "equipInstallCostF_VID": Designer.equipInstallCostF_VID,  # Installation Factor
            "equipMaintcCostF_VID": Designer.equipMaintcCostF_VID,  # Maintenance Factor
            "equipUsageRate_VID": Designer.equipUsageRate_VID,  # Usage Rate
            "equipAvailabilityRate_VID": Designer.equipAvailabilityRate_VID,  # Availability Rate
            "busyTime_VID": Designer.busyTime_VID,  # Busy Time
            "occupancyTime_VID": Designer.occupancyTime_VID,  # Occupancy Time
            "maxFillRatio_VID": Designer.maxFillRatio_VID,  # Maximum Fill Ratio
            "equipmentName_VID": Designer.equipmentName_VID,  # Equipment Name
            "description_VID": Designer.description_VID,  # Description
            "comments_VID": Designer.comments_VID,  # User Comments
            "size_VID": Designer.size_VID,  # Equipment Size
            "sizeUnits_VID": Designer.sizeUnits_VID,  # Equipment Size Units
            "sizeName_VID": Designer.sizeName_VID,  # Sizing Description
            "typeName_VID": Designer.typeName_VID,  # Equipment Type
            "typeID_VID": Designer.typeID_VID,  # Equipment Type ID
        }
        self.init_vars:dict[str,int] = {
            "autoInitMode_VID": Designer.autoInitMode_VID,              # Initialization Mode
            "isSourceFileSame_VID": Designer.isSourceFileSame_VID,      # Is source the same file?
            "bPartialPath_VID": Designer.bPartialPath_VID,              # Is Path Partial
            "sourceFileName_VID": Designer.sourceFileName_VID,          # Source File Name
            "sourceStreamName_VID": Designer.sourceStreamName_VID,      # Source Stream/Equipment Name
            "bInitTotalMassFlow_VID": Designer.bInitTotalMassFlow_VID,  # Initialize Total Mass Flow
            "totalFlowFactor_VID": Designer.totalFlowFactor_VID,        # Total Mass Flow Factor
            "bInitComposition_VID": Designer.bInitComposition_VID,      # Initialize Composition
            "bInitDensity_VID": Designer.bInitDensity_VID,              # Initialize Density / Thermo Data
            "bInitTemperature_VID": Designer.bInitTemperature_VID,      # Initialize Temperature
            "bInitPressure_VID": Designer.bInitPressure_VID,            # Initialize Pressure
            "bInitEntityData_VID": Designer.bInitEntityData_VID,        # Initialize Discrete Entity Data
            "solveAutoInitMode_VID": Designer.solveAutoInitMode_VID     # Before solving M&E balances auto initialization Mode
        }

        self.operation_vars:dict[str,int] = {
            "startTime_VID": Designer.startTime_VID,      # Operation Start Time
            "endTime_VID": Designer.endTime_VID,          # Operation End Time
            "setUpTime_VID": Designer.setUpTime_VID,      # Operation Set Up Time
            "processTime_VID": Designer.processTime_VID,  # Operation Process Time
            "holdupTime_VID": Designer.holdupTime_VID,    # Operation Holdup Time
            "turnaroundTime_VID": Designer.turnaroundTime_VID,  # Operation Turnaround Time
            "processTimeCalcMode_VID": Designer.processTimeCalcMode_VID,  # Process Time Calculation Mode
            "timeShift_VID": Designer.timeShift_VID,      # Operation Start Time Shift
            "thermalMode_VID": Designer.thermalMode_VID,  # Thermal Mode
            "exitTemperature_VID": Designer.exitTemperature_VID,  # Exit Temperature
            "heatingDuty_VID": Designer.heatingDuty_VID,  # Heating Duty
            "coolingDuty_VID": Designer.coolingDuty_VID,  # Cooling Duty
            "primaryHxAgentName_VID": Designer.primaryHxAgentName_VID,  # Primary Heat Transfer Agent Name
            "primaryHxAgentRate_VID": Designer.primaryHxAgentRate_VID,  # Primary Heat Transfer Agent Rate
            "primaryHxAgentDuty_VID": Designer.primaryHxAgentDuty_VID,  # Primary Heat Transfer Agent Duty
            "isPrimaryHxAgentHeating_VID": Designer.isPrimaryHxAgentHeating_VID,  # Is Primary Heat Transfer Agent Heating or Cooling?
            "powerCalcMode_VID": Designer.powerCalcMode_VID,  # Power Calculation Mode
            "power_VID": Designer.power_VID,              # Power
            "specPower_VID": Designer.specPower_VID,      # Specific Power
            "powerPerUnit_VID": Designer.powerPerUnit_VID,  # Power per Unit
            "powerDissipationFrac_VID": Designer.powerDissipationFrac_VID,  # Power Dissipation to Heat
            "opType_VID": Designer.opType_VID,            # Operation Type
            "opDescr_VID": Designer.opDescr_VID,          # Operation Description
            "isOpDescrSetByUser_VID": Designer.isOpDescrSetByUser_VID,  # Is Operation Description Set By User
            "auxEquipName_VID": Designer.auxEquipName_VID,  # Auxiliary Equipment Name used by Operation
            "comments_VID": Designer.comments_VID,        # User Comments
            "laborNeed_VID": Designer.laborNeed_VID,      # Labor Need
            "laborUnits_VID": Designer.laborUnits_VID,    # Labor Units
            "bIsVentOn_VID": Designer.bIsVentOn_VID,      # Is Vent On
            "emissionsFracs_VID": Designer.emissionsFracs_VID  # Emission Fractions
        }

        self.operations:dict[str,dict[str,int]] = {
            "generic_wash":{
                "operatingOption_VID": Designer.operatingOption_VID,    # Operating Option
                "volumeSpecOption_VID": Designer.volumeSpecOption_VID,  # Volume Specification Option
                "massBefore_VID": Designer.massBefore_VID,              # Contained Amount Before
                "removalFrac_VID": Designer.removalFrac_VID,            # Loss Fraction
                "massAfter_VID": Designer.massAfter_VID,                # Contained Amount After
                "washVolFlow_VID": Designer.washVolFlow_VID,            # Wash Volumetric Flowrate
                "washLossFrac_VID": Designer.washLossFrac_VID           # Wash Loss Fraction
            },
            "component_splitter":{
                "componentSplits_VID": Designer.componentSplits_VID,  # Component splits
                "bIsSplitToTopStream_VID": Designer.bIsSplitToTopStream_VID,  # Is splitting to top stream
                "bIsComponentSplit_VID": Designer.bIsComponentSplit_VID,  # Is component fraction split or flow split
                "componentFlow_VID": Designer.componentFlow_VID,  # Component Flow
            },
            "material_wash":{
                "flowrateOption_VID": Designer.flowrateOption_VID,   # Volume spec type
                "componentSplits_VID": Designer.componentSplits_VID, # Component fraction
                "washVolume_VID": Designer.washVolume_VID,           # Wash-in volume amount
                "relativeWashVolume_VID": Designer.relativeWashVolume_VID, # Relative wash-in volume amount
                "conversion_VID": Designer.conversion_VID,           # Equilibrium approach [0,1]
                "throughput_VID": Designer.throughput_VID,           # Throughput
                "temperature_VID": Designer.temperature_VID,         # Product stream temperature
                "isConversionSetByUser_VID": Designer.isConversionSetByUser_VID, # Is approach to equilibrium set by user?
                "bIgnoreEB_VID": Designer.bIgnoreEB_VID,             # Ignore energy balance?
                "retainedWashFrac_VID": Designer.retainedWashFrac_VID # Fraction of wash retained by the product
            }
        }



    #Application Related Methods:
    """
    These methods are used for performing general application tasks such as activating the designer, application, opening and closing files, etc.
    """
    def ShowApp(self):
        """
        ShowApp( ) This subroutine is used to activate the Pro-Designer application and display it in its current size position.
        """
        return self.app.ShowApp()

    def CloseApp(self):
        """
        CloseApp( ) This subroutine is used to close the Pro-Designer application. If there are Pro-Designer case files still open it will close all the documents without saving them.
        """
        return self.app.CloseApp()

    def OpenDoc(self,fileName:str):
        """
        OpenDoc(fileName As String) This function is used to open the Pro-Designer file with name fileName, makes this file the active Document object, and returns a reference to the caller.
        """
        self.doc = self.app.OpenDoc(fileName)
        return SuperProDesignerDocument(self,self.doc)

    def SetActiveDoc(self,fileName:str):
        """
        SetActiveDoc(fileName As String) This function is used to activate the Pro-Designer file with name fileName and also returns a reference to this file as a Document object.
        """
        raise NotImplementedError()
        self.doc = self.app.SetActiveDoc(fileName)
        print(self.doc)
        return SuperProDesignerDocument(self,self.doc)

    def CloseAllDocs(self,bSaveIfNeeded:bool):
        """
        CloseAllDocs(bSaveIfNeeded As Boolean) This subroutine is used to close all open Pro-Designer file (Document objects) Use bSaveIfNeeded = True for saving the Designer case files and bSaveIfNeeded = False for just closing the documents.
        """
        raise NotImplementedError()
        return self.app.CloseAllDocs(bSaveIfNeeded)

class SuperProDesignerDocument():
    def __init__(self,app:SuperProDesigner,doc):
        self.app = app
        self.doc = doc
    #Document Related Methods:
    """
    These methods are used for performing generic document tasks on specific Pro-Designer case files.
    """
    def CloseDoc(self,bSaveIfNeeded:bool):
        """
        CloseDoc(bSaveIfNeeded As Boolean) This subroutine is used to close the active Pro-Designer file (Document object). Use bSaveIfNeeded = True for saving the Designer case file and bSaveIfNeeded = False for just closing the document.
        """
        return self.doc.CloseDoc(bSaveIfNeeded)
    def GetDocName(self):
        raise NotImplementedError()
        """
        GetDocName(fileName As String, nMaxChar As Long) This function is used to return the name of the active Pro-Designer file (Document object). The function returns a Boolean which is True if it was successful in obtaining the file name and False if it was not. The filename argument is an output argument and returns the name of the Pro-Designer file. The nMaxChar is an input argument and specifies the number of characters that the file name will contain
        """
        doc_name = self.doc.GetDocName()
        return doc_name
    def SaveDoc(self):
        """
        SaveDoc() This subroutine is used to save the the active Document object.
        """
        return self.doc.SaveDoc()

    #Simulation Related Methods
    """
    These methods are used for simulation tasks. They are all functions that return a Boolean value, which is True if the task was successful and False if the task failed. They include:
    """
    def DoMEBalances(self):
        """
        DoMEBalances(val) This function is equivalent to clicking on the Solve button or to selecting Tasks / Do M&E Balances from the Pro-Designer application main menu. The value of variable (val) is currently of no importance. 
        """
        return self.doc.DoMEBalances(byref(VARIANT()))

    def DoEconomicCalculations(self):
        """
        DoEconomicCalculations( ) This function is equivalent to selecting Tasks / Perform Economic Calculations from the Pro-Designer application main menu. 
        """
        return self.doc.DoEconomicCalculations()
    def ScaleUpThroughput(self):
        raise NotImplementedError()
        """
        ScaleUpThroughput(VarID As VarID, val) This function is used for scaling the process throughput (It is equivalent to selecting Tasks / Adjust Process Throughput from the Pro-Designer application main menu and selecting the Based on Scale Up / Down Factor option). Use VarID = scaleUpFactor_VID and the value of the scale up factor for val (val is a Variant, it’s type should be double and its value should be greater than zero). 
        """
        return self.doc.ScaleUpThroughput()


    # Functions for Process (Flowsheet) Variables
    def RenameProcedure(self,oldName:str, newName:str):
        """
        RenameProcedure(oldName As String, newName As String)
        """
        return self.doc.RenameProcedure(oldName, newName)

    def RenameOperation(self, procedureName:str, oldName:str, newName:str):
        """
        RenameOperation(procedureName As String, oldName As String, newName As String)
        """
        return self.doc.RenameOperation(procedureName, oldName, newName)

    def RenameStream(self,oldName:str, newName:str):
        """
        RenameStream(oldName As String, newName As String)
        """
        return self.doc.RenameStream(oldName, newName)

    def RenameEquipment(self,oldName:str, newName:str):
        """
        RenameEquipment(oldName As String, newName As String)
        """
        return self.doc.RenameEquipment(oldName, newName)

    # Functions for Section Variables
    # Functions for Procedure Variables
    def GetUPVarVal(self, procName:str, VarID:int):
        """
        GetUPVarVal(procName As String, VarID As VarID, val)
        """
        out_var = VARIANT()
        if not self.doc.GetUPVarVal(procName,VarID,byref(out_var)): return False
        assert isinstance(out_var.value,(float,bool,str)),type(out_var.value)
        return out_var.value

    def GetUPVarVal2(self):
        """
        GetUPVarVal2(procName As String, VarID As VarID, val, val2)
        """
        raise NotImplementedError()
        return self.doc.GetUPVarVal2()

    def SetUPVarVal(self, procName:str, VarID:int, val):
        """
        SetUPVarVal(procName As String, VarID As VarID, val)
        """
        return self.doc.SetUPVarVal(procName,VarID,val)

    def SetUPVarVal2(self):
        """
        SetUPVarVal2(procName As String, VarID As VarID, val, val2)
        """
        raise NotImplementedError()
        return self.doc.SetUPVarVal2()

    def GetUPEmptiedContentsVarVal(self):
        """
        GetUPEmptiedContentsVarVal(procName As String, VarID As VarID, val, val2)
        """
        raise NotImplementedError()
        return self.doc.GetUPEmptiedContentsVarVal()


    # Functions for Equipment Variables
    def GetEquipVarVal(self,equipName:str,VarID:int):
        """
        GetEquipVarVal(equipName As String, VarID As VarID, val) 
        """
        out_var = VARIANT()
        if not self.doc.GetEquipVarVal(equipName,VarID,byref(out_var)): return False
        assert isinstance(out_var.value,(float,bool,str)),type(out_var.value)
        return out_var.value

    def GetEquipVarVal3(self):
        """
        GetEquipVarVal3(equipName As String, VarID As VarID, val, val2, val3)
        """
        raise NotImplementedError()
        return self.doc.GetEquipVarVal3()

    def SetEquipVarVal(self,equipName:str,VarID:int,val):
        """
        SetEquipVarVal(equipName As String, VarID As VarID, val)
        """
        return self.doc.SetEquipVarVal(equipName,VarID,val)

    def SetEquipVarVal3(self):
        """
        SetEquipVarVal3(equipName As String, VarID As VarID, val, val2, val3)
        """
        raise NotImplementedError()
        return self.doc.SetEquipVarVal3()

    def GetOperVarVal(self):
        """
        GetOperVarVal(procName As String, operName As String, VarID As VarID, val)
        """
        #GetOperVarVal2(procName As String, operName As String, VarID As VarID, val, val2)
        #GetOperVarVal3(procName As String, operName As String, VarID As VarID, val, val2, val3) 
        raise NotImplementedError
        return self.doc.GetOperVarVal()

    def SetOperVarVal(self, procName:str, operName:str, VarID:int, val, val2=None, val3=None):
        """
        SetOperVarVal(procName As String, operName As String, VarID As VarID, val)
        SetOperVarVal2(procName As String, operName As String, VarID As VarID, val, val2)
        SetOperVarVal3(procName As String, operName As String, VarID As VarID, val, val2, val3)
        """
        if val3 != None:return self.doc.SetOperVarVal3(procName, operName, VarID, val, val2, val3)
        if val2 != None:return self.doc.SetOperVarVal2(procName, operName, VarID, val, val2)
        if val != None:return self.doc.SetOperVarVal(procName, operName, VarID, val)
        return False


    # Functions for Stream Variables
    """
    The following functions can be used for setting or retrieving variables that refer to a specific stream (input /output /intermediate) that is included in the process file:
    """
    def GetStreamVarVal(self, streamName:str, VarID:int, compLocalName:str=''):
        """
        GetStreamVarVal(streamName As String, VarID As VarID, val, compLocalName As String) can be used to retrieve the value of input/output variables related to the specific stream
        VarID - Example "Designer.massFlow_VID"
        """
        out_var = VARIANT()
        if not self.doc.GetStreamVarVal(streamName, VarID, byref(out_var), compLocalName): return False
        assert isinstance(out_var.value,(float,bool)),type(out_var.value)
        return out_var.value

    def SetStreamVarVal(self, streamName:str, VarID:int, val:float|int, compLocalName:str='') -> bool:
        """
        SetStreamVarVal(streamName As String, VarID As VarID, val, compLocalName As String) can be used for setting input variables related to the specific stream
        VarID - Example "Designer.massFlow_VID"
        """
        assert isinstance(val,(float,int))
        return self.doc.SetStreamVarVal(streamName, VarID, float(val), compLocalName)

    def AddIngredientToInputStream(self, streamName:str, ingredientName:str, VarID:int, val:float):
        """
        AddIngredientToInputStream(streamName As String, ingredientName As String, VarID As VarID, val) can be used to add pure components and/or stock mixtures as well as the ingredient’s mass/mole flow or mass fraction to an input stream. The variable IDs that can be used with this function are: componentMassFlow_VID, componentMoleFlow_VID or compMassFrac_VID.
        """
        return self.doc.AddIngredientToInputStream(streamName, ingredientName, VarID, val)

    def RemoveIngredientFromInputStream(self,streamName:str,ingredientName:str):
        """
        RemoveIngredientFromInputStream(streamName As String, ingredientName As String) can be used to remove an ingredient from an input stream.
        """
        raise NotImplementedError()
        return self.doc.RemoveIngredientFromInputStream(streamName,ingredientName)

    def IsInputStreamCompositionValid(self,streamName:str):
        """
        IsInputStreamCompositionValid(streamName As String) can be used to validate the composition of an input stream. The sum of all ingredient mass fractions should add-up to 1.0.
        """
        return self.doc.IsInputStreamCompositionValid(streamName)


    # Functions for Ingredient Variables
    # Functions for Heat Transfer Agent Variables
    # Functions for Power Variables
    # Functions for Report Option Variables
    # Functions for Excel Data Link Variables
    # Functions for Excel Table Variables

    # Initialization
    def GetStreamAutoInitOptions(self):
        """
        GetStreamAutoInitOptions(streamName As String, varID As VarID, val) With this function you may retrieve all relevant setting that appear in the Input Stream Initialization Options Dialog.. The streamName is the name of the stream you wish to auto initialize. For a list of the Variable ID’s that can be used with this function , see Auto Initialization Variables.
        """
        raise NotImplementedError()
        return self.doc.GetStreamAutoInitOptions()

    def SetStreamAutoInitOptions(self,streamName:str,varID:int,val):
        """
        SetStreamAutoInitOptions(streamName As String, varID As VarID, val) With this function you may set all relevant setting that appear in the Input Stream Initialization Options Dialog. The streamName is the name of the stream <stream1> you wish to auto initialize by another stream <stream2> which you may also set in this function through the VARIANT val argument. For a list of the Variable ID’s that can be used with this function, see Auto Initialization Variables.
        """
        return self.doc.SetStreamAutoInitOptions(streamName,varID,val)

    def AutoInitStream(self,streamName:str):
        """
        AutoInitStream(streamName As String) This function is used to Automatically initialize the contents of the stream when the Auto-Initialize from Other Stream option has been checked, in the Input Stream Initialization Options Dialog.
        """
        return self.doc.AutoInitStream(streamName)

    def AutoInitAllStreamsAndEquipContents(self):
        """
        AutoInitAllStreamsAndEquipContents This function can also be used to auto initialize all the streams and equipment contents that have been set to be initialized by other sources.
        """
        raise NotImplementedError()
        return self.doc.AutoInitAllStreamsAndEquipContents()

    # Errors
    def GetCOMErrorMsg(self):
        """
        GetCOMErrorMsg(val) This function returns True if it was successful in obtaining the error message and False if it was not. The argument val (which must be defined as a Variant) contains the string with the error message.
        """
        out_var = VARIANT()
        if not self.doc.GetCOMErrorMsg(byref(out_var)): return False
        assert isinstance(out_var.value,str),type(out_var.value)
        return out_var.value
        
    def IsCOMSimDataComplete(self):
        """
        IsCOMSimDataComplete() This function returns  True if the data exchange using the COM functions was consistent. If the function returns False it means that some data that you have set using the COM functions is inconsistent and that you cannot proceed with the simulation. In this case you can use the GetCOMErrorMsg(val) to find out what went wrong.
        """
        raise NotImplementedError()
        return self.doc.IsCOMSimDataComplete()

    # Enumerator
    def Enumerator(self, ids:tuple[int,int], containerName1:str='',containerName2:str=''):
        pos = VARIANT()
        itemName = VARIANT()
        containerID,listID = ids

        if containerName2 == '':
            cont = self.doc.StartEnumeration(byref(pos),listID,containerID,containerName1)
            while cont:
                cont = self.doc.GetNextItemName(byref(pos),byref(itemName),listID,containerID,containerName1)
                yield itemName.value
        else:
            cont = self.doc.StartEnumeration2(byref(pos),listID,containerID,containerName1,containerName2)
            while cont:
                cont = self.doc.GetNextItemName2(byref(pos),byref(itemName),listID,containerID,containerName1,containerName2)
                yield itemName.value

    def Stream(self,initialName:str):
        return Stream(self,initialName)
    def Procedure(self,initialName:str):
        return Procedure(self,initialName)


class Stream:
    def __init__(self,doc:SuperProDesignerDocument,initialName:str):
        self.app = doc.app
        self.doc = doc
        self._name = initialName

        self.is_input_stream  = bool(self.doc.GetStreamVarVal(self.name,self.app.stream_vars["isInputStream_VID"]))
        self.is_output_stream = bool(self.doc.GetStreamVarVal(self.name,self.app.stream_vars["isOutputStream_VID"]))

    @property
    def name(self):
        return self._name
    
    @name.setter
    def name(self,newName:str):
        if not self.doc.RenameStream(self.name,newName): raise NameError("Stream name already exists")
        self._name = newName
 
    def AddIngredientToInputStream(self, ingredientName:str, VarID:int, val:float):
        """
        AddIngredientToInputStream(streamName As String, ingredientName As String, VarID As VarID, val) can be used to add pure components and/or stock mixtures as well as the ingredient’s mass/mole flow or mass fraction to an input stream. The variable IDs that can be used with this function are: componentMassFlow_VID, componentMoleFlow_VID or compMassFrac_VID.
        """
        if isinstance(val,int):val=float(val)
        return self.doc.AddIngredientToInputStream(self.name, ingredientName, VarID, val)

    def RemoveIngredientFromInputStream(self,streamName:str,ingredientName:str):
        """
        RemoveIngredientFromInputStream(streamName As String, ingredientName As String) can be used to remove an ingredient from an input stream.
        """
        raise NotImplementedError()
        return self.doc.RemoveIngredientFromInputStream(streamName,ingredientName)

    @property
    def IsInputStreamCompositionValid(self):
        assert self.is_input_stream
        return self.doc.IsInputStreamCompositionValid(self.name)

    @property
    def massFlow(self):
        return self.doc.GetStreamVarVal(self.name,self.app.stream_vars["massFlow_VID"])
    @massFlow.setter
    def massFlow(self,val:float):
        assert self.is_input_stream
        return self.doc.SetStreamVarVal(self.name,self.app.stream_vars["massFlow_VID"],val)

class Procedure:
    def __init__(self,doc:SuperProDesignerDocument,initialName:str):
        self.app = doc.app
        self.doc = doc
        self._name = initialName
        equipment_name = self.doc.GetUPVarVal(self.name,self.app.procedure_vars["equipmentName_VID"])
        assert isinstance(equipment_name,str)
        self.equipment = Equipment(self.doc,equipment_name)
        self._description = self.doc.GetUPVarVal(self.name,self.app.procedure_vars["description_VID"])

    @property
    def name(self):
        return self._name
    
    @name.setter
    def name(self,newName:str):
        if not self.doc.RenameProcedure(self.name,newName): raise NameError("Procedure name already exists")
        self._name = newName

    @property
    def description(self):
        return self._description
    
    @description.setter
    def description(self,newDescription:str):
        if not self.doc.SetUPVarVal(self.name,self.app.procedure_vars["description_VID"],newDescription): raise NameError("Desription failed to assign?")
        self._description = newDescription

    def Operation(self,initialName:str):
        return Operation(self,initialName)

class Equipment:
    def __init__(self,doc:SuperProDesignerDocument,initialName:str):
        self.app = doc.app
        self.doc = doc
        self._name = initialName
        self._description = self.doc.GetEquipVarVal(self.name,self.app.equipment_vars["description_VID"])

    @property
    def name(self):
        return self._name
    
    @name.setter
    def name(self,newName:str):
        if not self.doc.RenameEquipment(self.name,newName): raise NameError("Equipment name already exists")
        self._name = newName

    @property
    def description(self):
        return self._description
    
    @description.setter
    def description(self,newDescription:str):
        if not self.doc.SetEquipVarVal(self.name,self.app.equipment_vars["description_VID"],newDescription): raise NameError("Desription failed to assign?")
        self._description = newDescription

class Operation:
    def __init__(self,proc:Procedure,initialName:str):
        self.app = proc.doc.app
        self.doc = proc.doc
        self.proc = proc
        self._name = initialName

    @property
    def name(self):
        return self._name
    
    @name.setter
    def name(self,newName:str):
        if not self.doc.RenameOperation(self.proc.name,self.name,newName): raise NameError("Operation name already exists")
        self._name = newName

    def SetOperVarVal(self, VarID:int, val, val2=None, val3=None):
        return self.doc.SetOperVarVal(self.proc.name, self.name, VarID, val, val2, val3)

if __name__ == "__main__":
    spd = SuperProDesigner()
    spd.ShowApp()
    doc = spd.OpenDoc(r"C:\Users\mikbr\Desktop\SPD\COM\v1.spf")

    print(doc.GetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,""))
    print(doc.SetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,10,""))
    print(doc.GetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,""))
    print(doc.SetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,20,""))
    print(doc.GetStreamVarVal("Dirty chokeberries",spd.Designer.massFlow_VID,""))
    print(doc.DoMEBalances())

    for stream in doc.Enumerator(spd.enum_vars["flowsheet_CID"]["stream_LID"]):
        print(stream)

    embed()

    input()
    doc.CloseDoc(False)
    input()
    spd.CloseApp()