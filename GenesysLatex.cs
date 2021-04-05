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

namespace genesys_latex
{
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
                // Tables
                Losses(project, outputPath);
                Hazards(project, outputPath);

                // Diagrams
                Architecture(project, outputPath);
                client.Dispose();
                Console.WriteLine("Close Connection");
                System.Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Usage: <project name> <output path>");
            }

        }
        static void Losses( IProject project, string outputPath)
        {
            IFolder categoryFolder = project.GetFolder("Category");
            IEntity systemLossCategory = categoryFolder.GetEntity("STPA: Loss");

            IEnumerable<IEntity> lossList = systemLossCategory.GetRelationshipTargets("categorizes");
            ISortBlock lossSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock hazardSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

            // Write Latex Table Definition Header
            dataFile = new System.IO.StreamWriter(outputPath + @"tables/loss.tex", false);
            dataFile.WriteLine(@"% ++++++++++++++++");
            dataFile.WriteLine(@"% STPA Loss Table");
            dataFile.WriteLine(@"% ++++++++++++++++");

            dataFile.WriteLine(@"\begin{table}[ht]");
            dataFile.WriteLine(@"\scriptsize");
            dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
            dataFile.WriteLine(@"\centering");
            dataFile.WriteLine(@"\caption{" + project.Name + " STPA Losses}");
            dataFile.WriteLine(@"\begin{tabular}{l c l}");
            dataFile.WriteLine(@"    \toprule");
            dataFile.WriteLine(@"    \textbf{Loss} & \textbf{Priority} & \textbf{is caused by: Hazard} \\");


            foreach (IEntity loss in lossSortBlock.SortEntities(lossList))
            {
                dataFile.WriteLine(@"    \midrule");

                var outLine = string.Format(@"    {0}:{1}&{2}&",
                        loss.GetAttribute("number").ToString(), loss.Name,
                        loss.GetAttribute("priority").ToString());

                IEnumerable<IEntity> hazardList = loss.GetRelationshipTargets("is caused by");
                int index = 0;
                foreach (IEntity hazard in hazardSortBlock.SortEntities(hazardList))
                {
                    outLine += Environment.NewLine + @"    ";
                    if (index > 0)
                    {
                        outLine += @"& &";
                    }
                    outLine += string.Format(@"{0}:{1}\\", hazard.GetAttribute("number").ToString(), hazard.Name);
                    index++;
                }
                dataFile.WriteLine(outLine);
            }
            dataFile.WriteLine(@"\bottomrule");
            dataFile.WriteLine(@"\end{tabular}");
            dataFile.WriteLine(@"\label{table:stpa-losses}");
            dataFile.WriteLine(@"\end{table}");

            dataFile.Close();
        }
        static void Hazards(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity systemHazardCategory = cateogoryFolder.GetEntity("STPA: Hazard");

            IEnumerable<IEntity> hazardList = systemHazardCategory.GetRelationshipTargets("categorizes");
            ISortBlock lossSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock hazardSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock hcaSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

            // Write Latex Table Definition Header
            dataFile = new System.IO.StreamWriter(outputPath + @"tables/hazard.tex", false);
            dataFile.WriteLine(@"% ++++++++++++++++++");
            dataFile.WriteLine(@"% STPA Hazard Table");
            dataFile.WriteLine(@"% ++++++++++++++++++");

            dataFile.WriteLine(@"\begin{table}[ht]");
            dataFile.WriteLine(@"\scriptsize");
            dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
            dataFile.WriteLine(@"\centering");
            dataFile.WriteLine(@"\caption{" + project.Name + " STPA Hazards}");
            dataFile.WriteLine(@"\begin{tabular}{l l l}");
            dataFile.WriteLine(@"    \toprule");
            dataFile.WriteLine(@"    \textbf{Hazard} & \textbf{leads to: Loss} & \textbf{is caused by: Hazardous Control Action} \\");


            foreach (IEntity hazard in hazardSortBlock.SortEntities(hazardList))
            {
                dataFile.WriteLine(@"    \midrule");

                var outLine = string.Format(@"    {0}:{1}&",
                        hazard.GetAttribute("number").ToString(), hazard.Name);

                ListCollectionView hcaList = hcaSortBlock.SortEntities(hazard.GetRelationshipTargets("is caused by"));
                hcaList.MoveCurrentToFirst();
                ListCollectionView lossList = lossSortBlock.SortEntities(hazard.GetRelationshipTargets("leads to"));
                lossList.MoveCurrentToFirst();
                var maxItems = Math.Max(hcaList.Count, lossList.Count);
                for (var count = 1; count <= maxItems; count++)
                {
                    outLine += Environment.NewLine + @"    ";
                    if (count > 1)
                    {
                        outLine += @"&";
                    }
                    if (lossList.Count >= count)
                    {
                        IEntity loss = (IEntity)lossList.CurrentItem;
                        if (hcaList.Count >= count)
                        {
                            IEntity hca = (IEntity)hcaList.CurrentItem;
                            outLine += string.Format(@"{0}:{1}&{2}:{3}\\", 
                                loss.GetAttribute("number").ToString(), loss.Name,
                                hca.GetAttribute("number").ToString(), hca.Name);
                            hcaList.MoveCurrentToNext();
                        }
                        else
                        {
                            outLine += string.Format(@"{0}:{1}&\\",
                                loss.GetAttribute("number").ToString(), loss.Name);
                        }
                        lossList.MoveCurrentToNext();
                    }
                    else
                    {
                        IEntity hca = (IEntity)hcaList.CurrentItem;
                        outLine += string.Format(@"&{0}:{1}\\",
                            hca.GetAttribute("number").ToString(), hca.Name);
                        hcaList.MoveCurrentToNext();
                    }
                }
                dataFile.WriteLine(outLine);
            }
            dataFile.WriteLine(@"\bottomrule");
            dataFile.WriteLine(@"\end{tabular}");
            dataFile.WriteLine(@"\label{table:stpa-hazards}");
            dataFile.WriteLine(@"\end{table}");

            dataFile.Close();
        }
        static void Architecture(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity systemArchitectureCategory = cateogoryFolder.GetEntity("CSRM: Architecture");

            IEnumerable<IEntity> componentList = systemArchitectureCategory.GetRelationshipTargets("categorizes");

            foreach (IEntity component in componentList)
            {
                DiagramExporter diagramExporter = new DiagramExporter(component,
                    EntityDiagramType.PhysicalBlock, null)
                {
                    Background = Colors.White
                };
                String componentDiagramName = component.Name.Replace(@":", "_").Replace(" ", "_") + @"_" +
                    EntityDiagramType.PhysicalBlock.ToString();

                String diagramFilePath = String.Format("{0}diagrams/{1}.png", outputPath, componentDiagramName);

                using (MemoryStream stream = new MemoryStream())
                {
                    diagramExporter.Export(stream);
                    using (FileStream file = new FileStream(diagramFilePath, FileMode.Create, FileAccess.Write))
                    {
                        stream.WriteTo(file);
                    }
                }
                String diagramFigurePath = String.Format("{0}diagrams/{1}.tex", outputPath, componentDiagramName);

                // Write Latex Figure Definition Header
                dataFile = new System.IO.StreamWriter(diagramFigurePath, false);
                dataFile.WriteLine(@"\begin{figure}");
                dataFile.WriteLine(@"    \centering");
                dataFile.WriteLine(@"    \includegraphics[width =\linewidth]{mbse/diagrams/" + 
                    componentDiagramName + @".png}");
                dataFile.WriteLine(@"    \caption{" + project.Name + @": " + 
                    component.Name.Replace(@"_", @" ") + @": " + 
                    EntityDiagramType.PhysicalBlock.ToString() + @"}");
                dataFile.WriteLine(@"    \label{fig:" + componentDiagramName + @"}");
                dataFile.WriteLine(@"\end{figure}");
                dataFile.Close();
            }
        }
    }
}
