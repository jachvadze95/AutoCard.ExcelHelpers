using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Reflection;
using System.Collections.Generic;
using System;
using System.Linq;
using AutoCard.ExcelHelpers.Services;
using AutoCard.ExcelHelpers.Models;
using System.Runtime.InteropServices;

namespace AutoCard.ExcelHelpers
{
    public class LineBackup
    {
        public ObjectId LineId { get; set; }
        public int ColorIndex { get; set; }
    }

    public class MyPlugin : IExtensionApplication
    {
        private Room room = new Room();
        private List<LineBackup> _recordedLines = new List<LineBackup>();


        private readonly Document doc;
        private readonly Database db;
        private readonly Editor ed;

        public MyPlugin()
        {
            doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            db = doc.Database;
            ed = doc.Editor;
        }

        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        [CommandMethod("CalculateRoomWalls")]
        public void CalculateRoomWalls()
        {
            bool kill = false;
            int iteration = 0;

            room.RoomNo = PromptVaraibleAssign<int>("Add RoomNo");
            room.RoomName = PromptVaraibleAssign<string>("Add RoomName");
            room.Height = PromptVaraibleAssign<double>("Add RoomHeight");

            while (!kill)
            {
                //After first iteration We ask for termination
                if (iteration > 0)
                {
                    var res = PromptVaraibleAssign<string>("Continue? Y/n");
                    if (res == "n")
                    {
                        kill = true;
                    }
                }


                var objectId = SelectLine();
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Line line = tr.GetObject(objectId, OpenMode.ForWrite) as Line;

                    if (line != null && !_recordedLines.Any(x => x.LineId == line.ObjectId))
                    {
                        AddWall(line);
                        line.ColorIndex = 1;
                    }

                    tr.Commit();
                }

                iteration++;
            }


            UpdateInExcel();

        }

        private ObjectId SelectLine()
        {
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect a line: ");
            peo.SetRejectMessage("Selected entity is not a line.");
            peo.AddAllowedClass(typeof(Line), false);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK) throw new ProgramExitExeption();

            return per.ObjectId;
        }

        private T PromptVaraibleAssign<T>(string description)
        {
            object res = null;
            bool correctPropValue = false;
            Type propType = typeof(T);

            while (!correctPropValue)
            {
                bool failed = false;

                var prompt = new PromptStringOptions(description);
                var result = ed.GetString(prompt);
                if (result.Status != PromptStatus.OK) throw new ProgramExitExeption();

                res = StaticHelpers.ConvertStringToType(result.StringResult, propType);
                if (res == null) failed = true;

                if (failed)
                {
                    ed.WriteMessage($"\nInput Corret Parameter Fomat for {propType}");
                }
                else
                {
                    correctPropValue = true;
                }
            }

            return (T)res;
        }

        private void UpdateInExcel()
        {
            using (IExcelService service = new ExcelService())
            {
                service.WriteDataWithBackup(room);
                ed.WriteMessage("\nWalls Updated In Excel");
                ResetHighlight();
            }
        }

        private void AddWall(Line line)
        {
            RoomWall wall = new RoomWall()
            {
                Length = line.Length,
                WallDirection = PromptVaraibleAssign<DirectionType>("Select Direction")
            };

            room.Walls.Add(wall);
            
            _recordedLines.Add(new LineBackup
            {
                ColorIndex = line.ColorIndex,
                LineId = line.ObjectId
            });

            ed.WriteMessage($"\nWall Added");
        }

        private void ResetHighlight()
        {
            if (!_recordedLines.Any()) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                foreach (var lineBackup in _recordedLines)
                {
                    Line line = tr.GetObject(lineBackup.LineId, OpenMode.ForWrite) as Line;

                    if (line != null)
                    {
                        line.ColorIndex = lineBackup.ColorIndex;
                    }

                }

                tr.Commit();
            }

            _recordedLines.Clear();
        }
    }
}
