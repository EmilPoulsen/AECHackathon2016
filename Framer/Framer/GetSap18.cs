using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Grasshopper;
using GH_IO;
using Rhino.Geometry;
using Grasshopper.Kernel;
using SAP2000v18;

namespace Framer
{
    public class GetSap18 : GH_Component
    {
        public GetSap18() : base("getSap18", "getSap18", "getSap18", "Framer", "Framer"){
        }

        public override GH_Exposure Exposure {
            get {
                return GH_Exposure.primary;
            }
        }

        public override Guid ComponentGuid {
            get {
                return new Guid("fbcd93af-0e4f-46cb-beba-1d5adcd99b4b");
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager) {
            pManager.AddBooleanParameter("run", "run", "run", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager) {
            pManager.AddGenericParameter("sap", "sap", "sap", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA) {
            bool run = false;
            cSapModel sapModel = null;

            if (!DA.GetData<bool>(0, ref run)) { return; }

            if (run) {

                cOAPI mySapObject = null;
                try {

                    //get the active Sap2000 object

                    mySapObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");

                } catch (Exception ex) {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "SAP instance not found");
                    //Console.WriteLine("No running instance of the program found or failed to attach.");
                    return;

                }

                sapModel = mySapObject.SapModel;
            }

            DA.SetData(0, sapModel);

            /*

                If run Then

      ' - attach to running SAP - - - - - - - - - - - - - - - - - - - - - - - - -
      'BV_Model10_R8.sdb

      Dim mySapObject As cOAPI
      mySapObject = Nothing

      'attach to a running instance of Sap2000
      'get the active Sap2000 object
      mySapObject = GetObject(, "CSI.SAP2000.API.SapObject")
      AppActivate("SAP2000")

      'Get a reference to cSapModel to access all OAPI classes and functions
      'Dim mySapModel As cSapModel
      mySapModel = mySapObject.SapModel

    End If


    SapModel_ = mySapModel
    */

        }
    }
}
