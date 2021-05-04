using System;
using System.Collections.Generic;
using System.Windows.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Vitech.Genesys.Client;
using Vitech.Genesys.Common;
using Vitech.Genesys.Client.Schema;
using Vitech.Genesys.Client.Simulation;
using Vitech.Genesys.Windows.Diagrams;
using System.IO;
using System.Windows.Media;
using Newtonsoft.Json;


namespace genesys_latex
{
    enum DiagramScale
    {
        LineWidth,
        ThreeQuarter,
        Half,
        Quarter
    }
    class GenesysLatex
    {
        static ClientModel client;
        static System.IO.StreamWriter dataFile;
        [STAThread]
        static void Main(string[] args)
        {
            String projectName;
            String outputPath;

            if (args.Length == 2) 
            {
                projectName = args[0];
                outputPath = args[1];
                // Setup Connection to GENESYS
                client = new ClientModel();
                RepositoryConfiguration repositoryConfiguration = client.GetKnownRepositories().LocalRepository;
                GenesysClientCredentials credentials = new GenesysClientCredentials("api-user", "api-pwd", AuthenticationType.GENESYS);
                repositoryConfiguration.Login(credentials);
                Console.WriteLine("Logged In!");

                Repository repository = repositoryConfiguration.GetRepository();
                IProject project = repository.GetProject(projectName);

                // Process JSON input for Tables and Diagrams
                CreateTables(project, outputPath);
                CreateDiagrams(project, outputPath);
                             
                client.Dispose();
                Console.WriteLine("Close Connection");
                System.Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Usage: <project name> <output path>");
            }
        }
        public class Diagram
        {
            public string GnsxCategory { get; set; }
            public string DiagramType { get; set; }
            public string Scale { get; set; }
        }
        public class Column
        {
            public string Heading { get; set; }
            public string GnsxType { get; set; }
            public string GnsxName { get; set; }
        }
        public class Table
        {
            public string Name { get; set; }
            public string File { get; set; }
            public string Caption { get; set; }
            public string Label { get; set; }
            public string GnsxCategory { get; set; }
            public string ColumnFormat { get; set; }
            public string SelfColumnHeading { get; set; }
            public IList<Column> Columns { get; set; }
        }
        static string EscapeString(string input)
        {
            return (input.
                Replace(@"\ ", @"\textbackslash \space ").
                Replace(@"\", @"\textbackslash ").
                Replace(@"&", @"\&").
                Replace(@"%", @"\%").
                Replace(@"$", @"\$").
                Replace(@"#", @"\#").
                Replace(@"_", @"\_").
                Replace(@"{", @"\{").
                Replace(@"}", @"\}").
                Replace(@"~ ", @"\textasciitilde \space ").
                Replace(@"~", @"\textasciitilde ").
                Replace(@"^ ", @"\textasciicircum \space ").
                Replace(@"^", @"\textasciicircum "));

        }
        static void CreateTables(IProject project, string outputPath)
        {
            IList<Table> tables = JsonConvert.DeserializeObject<List<Table>>(File.ReadAllText(@"..\..\tables.json"));
            foreach (Table table in tables)
            {
                IFolder cateogoryFolder = project.GetFolder("Category");
                IEntity tableCategory = cateogoryFolder.GetEntity(table.GnsxCategory);

                IEnumerable<IEntity> entityList = tableCategory.GetRelationshipTargets("categorizes");

                ISortBlock tableSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

                // up to 3 related entity columns per table definition
                List<ISortBlock> colSortBlocks = new List<ISortBlock>
                {
                    project.GetSortBlock(SortBlockConstants.Numeric),
                    project.GetSortBlock(SortBlockConstants.Numeric),
                    project.GetSortBlock(SortBlockConstants.Numeric)
                };

                List<ListCollectionView> colLists = new List<ListCollectionView>
                {
                    null,
                    null,
                    null
                };


                // Write Latex Table Definition Header
                dataFile = new System.IO.StreamWriter(outputPath + @"tables/" + table.File, false);
                dataFile.WriteLine(@"% " + new String('+', table.Name.Length));
                dataFile.WriteLine(@"% " + table.Name);
                dataFile.WriteLine(@"% " + new String('+', table.Name.Length));
                dataFile.WriteLine(@"\begin{table}[ht]");
                dataFile.WriteLine(@"\scriptsize");
                dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
                dataFile.WriteLine(@"\centering");
                dataFile.WriteLine(@"\caption{" + project.Name + " " + table.Caption + "}");
                dataFile.WriteLine(@"\begin{tabular}{" + table.ColumnFormat + "}");
                dataFile.WriteLine(@"    \toprule");

                // Write Table Headings
                string headingLine = @"    \textbf{" + table.SelfColumnHeading + "}";
                foreach (Column column in table.Columns )
                {
                    headingLine += @" & \textbf{" + column.Heading + "}";
                }
                headingLine += @"\\";
                dataFile.WriteLine(headingLine);

                // Write Table Rows
                foreach (IEntity entity in tableSortBlock.SortEntities(entityList))
                {
                    dataFile.WriteLine(@"    \midrule");

                    var outLine = string.Format(@"    {{{0}:{1}}}",
                            entity.GetAttribute("number").ToString(), EscapeString(entity.Name));

                    foreach (Column column in table.Columns)
                    {
                        if (column.GnsxType == "attribute")
                        {
                            outLine += Environment.NewLine + @"    &" 
                                        + entity.GetAttribute(column.GnsxName).ToString();
                        }
                    }
                    int numRelCol = 0;
                    foreach (Column column in table.Columns)
                    {
                        if (column.GnsxType == "relation")
                        {
                            colLists[numRelCol] = colSortBlocks[numRelCol].SortEntities(entity.GetRelationshipTargets(column.GnsxName));
                            colLists[numRelCol].MoveCurrentToFirst();
                            numRelCol++;
                        }
                    }

                    var maxItems = Math.Max(colLists[0]?.Count ?? 0, 
                        Math.Max(colLists[1]?.Count ?? 0, colLists[2]?.Count ?? 0));

                    if (maxItems > 0)
                    {   // seperator between entity and attributes (if exist)
                        outLine += @"&";
                    }
                    for (var count = 1; count <= maxItems; count++)
                    {
                        outLine += Environment.NewLine + @"    ";
                        if (count > 1)
                        {
                            // leave blanks for entity id/name and attributes
                            outLine += new String('&', table.Columns.Count + 1 - numRelCol);
                        }
                        if ((colLists[0]?.Count ?? 0) >= count)
                        {
                            IEntity col0Entity = (IEntity)colLists[0].CurrentItem;
                            if ((colLists[1]?.Count ?? 0) >= count)
                            {
                                IEntity col1Entity = (IEntity)colLists[1].CurrentItem;
                                if ((colLists[2]?.Count ?? 0) >= count)
                                {
                                    IEntity col2Entity = (IEntity)colLists[2].CurrentItem;
                                    outLine += string.Format(@"{0}:{1}&{2}:{3}&{4}:{5}\\",
                                        col0Entity.GetAttribute("number").ToString(), col0Entity.Name,
                                        col1Entity.GetAttribute("number").ToString(), col1Entity.Name,
                                        col2Entity.GetAttribute("number").ToString(), col2Entity.Name);

                                    colLists[2].MoveCurrentToNext();
                                }
                                else
                                {
                                    outLine += string.Format(@"{0}:{1}&{2}:{3}",
                                        col0Entity.GetAttribute("number").ToString(), col0Entity.Name,
                                        col1Entity.GetAttribute("number").ToString(), col1Entity.Name);
                                    outLine += new String('&', numRelCol - 2) + @"\\";

                                }
                                colLists[1].MoveCurrentToNext();
                            }
                            else
                            {
                                if ((colLists[2]?.Count ?? 0) >= count)
                                {
                                    IEntity col2Entity = (IEntity)colLists[2].CurrentItem;
                                    outLine += string.Format(@"{0}:{1}&&{2}:{3}\\",
                                        col0Entity.GetAttribute("number").ToString(), col0Entity.Name,
                                        col2Entity.GetAttribute("number").ToString(), col2Entity.Name);

                                    colLists[2].MoveCurrentToNext();
                                }
                                else
                                {
                                    outLine += string.Format(@"{0}:{1}",
                                        col0Entity.GetAttribute("number").ToString(), col0Entity.Name);
                                    outLine += new String('&', numRelCol - 1) + @"\\";
                                }
                            }
                            colLists[0].MoveCurrentToNext();
                        }
                        else
                        {
                            if ((colLists[1]?.Count ?? 0) >= count)
                            {
                                IEntity col1Entity = (IEntity)colLists[1].CurrentItem;
                                if ((colLists[2]?.Count ?? 0)>= count)
                                {
                                    IEntity col2Entity = (IEntity)colLists[2].CurrentItem;
                                    outLine += string.Format(@"&{0}:{1}&{2}:{3}\\",
                                        col1Entity.GetAttribute("number").ToString(), col1Entity.Name,
                                        col2Entity.GetAttribute("number").ToString(), col2Entity.Name);

                                    colLists[2].MoveCurrentToNext();
                                }
                                else
                                {
                                    outLine += string.Format(@"&{0}:{1}",
                                        col1Entity.GetAttribute("number").ToString(), col1Entity.Name);
                                    outLine += new String('&', numRelCol - 2) + @"\\";
                                }
                                colLists[1].MoveCurrentToNext();
                            }
                            else
                            {
                                if ((colLists[2]?.Count ?? 0) >= count)
                                {
                                    IEntity col2Entity = (IEntity)colLists[2].CurrentItem;
                                    outLine += string.Format(@"&&{0}:{1}\\",
                                        col2Entity.GetAttribute("number").ToString(), col2Entity.Name);

                                    colLists[2].MoveCurrentToNext();
                                }
                            }
                        }
                    }
                    if (maxItems == 0)
                    {
                        // no related instances - leave blank columns
                        outLine += new String('&', numRelCol) + @"\\";
                    }
                    dataFile.WriteLine(outLine);
                }

                // Write Table Footer
                dataFile.WriteLine(@"    \bottomrule");
                dataFile.WriteLine(@"\end{tabular}");
                dataFile.WriteLine(@"\label{" + table.Label + "}");
                dataFile.WriteLine(@"\end{table}");
                dataFile.Close();
            }
        }

        static void CreateDiagrams(IProject project, string outputPath)
        {
            IList<Diagram> diagrams = JsonConvert.DeserializeObject<List<Diagram>>(File.ReadAllText(@"..\..\figures.json"));
            foreach (Diagram diagram in diagrams)
            {
                IFolder cateogoryFolder = project.GetFolder("Category");
                IEntity diagramCategory = cateogoryFolder.GetEntity(diagram.GnsxCategory);

                IEnumerable<IEntity> entityList = diagramCategory.GetRelationshipTargets("categorizes");

                Enum.TryParse(diagram.DiagramType, out EntityDiagramType type);
                Enum.TryParse(diagram.Scale, out DiagramScale scale);

                foreach (IEntity entity in entityList)
                {
                    OutputDiagram(project, type, scale, entity, outputPath);
                }
            }
        }
        
        static void OutputDiagram(IProject project, EntityDiagramType diagramType, DiagramScale scale,
            IEntity diagramEntity, string outputPath)
        {
            IStoredView view = diagramEntity.GetDefaultStoredView(diagramType, null);
            DiagramExporter diagramExporter = new DiagramExporter(view)
            {
                Background = Colors.White
            };
            String diagramName = diagramEntity.Name.Replace(@":", "_").Replace(" ", "_") + @"_" +
                diagramType.ToString();

            String diagramFilePath = String.Format("{0}diagrams/{1}.png", outputPath, diagramName);

            using (MemoryStream stream = new MemoryStream())
            {
                diagramExporter.Export(stream);
                using (FileStream file = new FileStream(diagramFilePath, FileMode.Create, FileAccess.Write))
                {
                    stream.WriteTo(file);
                }
            }
            String diagramFigurePath = String.Format("{0}diagrams/{1}.tex", outputPath, diagramName);

            // Write Latex Figure Definition Header
            dataFile = new System.IO.StreamWriter(diagramFigurePath, false);
            dataFile.WriteLine(@"\begin{figure}");
            dataFile.WriteLine(@"    \centering");
            switch (scale)
            {
                case DiagramScale.LineWidth:
                    dataFile.WriteLine(@"    \includegraphics[width=\linewidth]{mbse/diagrams/" +
                        diagramName + @".png}");
                    break;
                case DiagramScale.ThreeQuarter:
                    dataFile.WriteLine(@"    \includegraphics[scale=.75]{mbse/diagrams/" +
                        diagramName + @".png}");
                    break;
            }

            dataFile.WriteLine(@"    \caption{" + project.Name + @": " +
                diagramEntity.Name.Replace(@"_", @" ") + @": " +
                diagramType.ToString() + @"}");
            dataFile.WriteLine(@"    \label{fig:" + diagramName + @"}");
            dataFile.WriteLine(@"\end{figure}");
            dataFile.Close();
        }
    }
}
