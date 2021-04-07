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
                
                // CSRM #1 - System Description
                SystemDescription(project, outputPath);
                // CSRM #2 - Operational Risk Assessment
                OperationalRisk(project, outputPath);
                // CSRM #3 - Resilient Modes
                ResilientModes(project, outputPath);
                // CSRM #4 - Cyber Vulnerability Assessment
                Vulnerability(project, outputPath);
                client.Dispose();
                Console.WriteLine("Close Connection");
                System.Environment.Exit(0);
            }
            else
            {
                Console.WriteLine("Usage: <project name> <output path>");
            }
        }
        static void Vulnerability(IProject project, string outputPath)
        {
            LossScenarios(project, outputPath);
            Remediations(project, outputPath);
        }
        static void LossScenarios(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity lossScenarioCategory = cateogoryFolder.GetEntity("STPA: Loss Scenario");

            IEnumerable<IEntity> lossScenarioList = lossScenarioCategory.GetRelationshipTargets("categorizes");
            ISortBlock lossScenarioSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock hcaSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock resilientModeSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

            // Write Latex Table Definition Header
            dataFile = new System.IO.StreamWriter(outputPath + @"tables/lossscenario.tex", false);
            dataFile.WriteLine(@"% ++++++++++++++++++++++++");
            dataFile.WriteLine(@"% STPA Loss Scenario Table");
            dataFile.WriteLine(@"% ++++++++++++++++++++++++");

            dataFile.WriteLine(@"\begin{table}[ht]");
            dataFile.WriteLine(@"\scriptsize");
            dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
            dataFile.WriteLine(@"\centering");
            dataFile.WriteLine(@"\caption{" + project.Name + " STPA Loss Scenarios}");
            dataFile.WriteLine(@"\begin{tabular}{l l l}");
            dataFile.WriteLine(@"    \toprule");
            dataFile.WriteLine(@"    \textbf{Loss Scenario} & \textbf{leads to: Hazardous Action} & \textbf{reconfigures using: Resilient Mode} \\");


            foreach (IEntity lossScenario in lossScenarioSortBlock.SortEntities(lossScenarioList))
            {
                dataFile.WriteLine(@"    \midrule");

                var outLine = string.Format(@"    {0}:{1}&",
                        lossScenario.GetAttribute("number").ToString(), lossScenario.Name);

                ListCollectionView hcaList = hcaSortBlock.SortEntities(lossScenario.GetRelationshipTargets("leads to"));
                hcaList.MoveCurrentToFirst();
                ListCollectionView resilientModeList = resilientModeSortBlock.SortEntities(lossScenario.GetRelationshipTargets("reconfigures using"));
                resilientModeList.MoveCurrentToFirst();
                var maxItems = Math.Max(hcaList.Count, resilientModeList.Count);
                for (var count = 1; count <= maxItems; count++)
                {
                    outLine += Environment.NewLine + @"    ";
                    if (count > 1)
                    {
                        outLine += @"&";
                    }
                    if (hcaList.Count >= count)
                    {
                        IEntity hca = (IEntity)hcaList.CurrentItem;
                        if (resilientModeList.Count >= count)
                        {
                            IEntity resilientMode = (IEntity)resilientModeList.CurrentItem;
                            outLine += string.Format(@"{0}:{1}&{2}:{3}\\",
                                hca.GetAttribute("number").ToString(), hca.Name,
                                resilientMode.GetAttribute("number").ToString(), resilientMode.Name);
                            resilientModeList.MoveCurrentToNext();
                        }
                        else
                        {
                            outLine += string.Format(@"{0}:{1}&\\",
                                hca.GetAttribute("number").ToString(), hca.Name);
                        }
                        hcaList.MoveCurrentToNext();
                    }
                    else
                    {
                        IEntity resilientMode = (IEntity)resilientModeList.CurrentItem;
                        outLine += string.Format(@"&{0}:{1}\\",
                            resilientMode.GetAttribute("number").ToString(), resilientMode.Name);
                        resilientModeList.MoveCurrentToNext();
                    }
                }
                dataFile.WriteLine(outLine);
            }
            dataFile.WriteLine(@"\bottomrule");
            dataFile.WriteLine(@"\end{tabular}");
            dataFile.WriteLine(@"\label{table:stpa-lossscenarios}");
            dataFile.WriteLine(@"\end{table}");

            dataFile.Close();
        }
        static void Remediations(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity remediationCategory = cateogoryFolder.GetEntity("CSRM: Remediations");

            IEnumerable<IEntity> remediationList = remediationCategory.GetRelationshipTargets("categorizes");
            ISortBlock remediationSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock implementationSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock protectionSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

            // Write Latex Table Definition Header
            dataFile = new System.IO.StreamWriter(outputPath + @"tables/remediation.tex", false);
            dataFile.WriteLine(@"% ++++++++++++++++++++++");
            dataFile.WriteLine(@"% CSRM Remediation Table");
            dataFile.WriteLine(@"% ++++++++++++++++++++++");

            dataFile.WriteLine(@"\begin{table}[ht]");
            dataFile.WriteLine(@"\scriptsize");
            dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
            dataFile.WriteLine(@"\centering");
            dataFile.WriteLine(@"\caption{" + project.Name + " CSRM Remediations}");
            dataFile.WriteLine(@"\begin{tabular}{p{0.3\textwidth} p{0.3\textwidth} p{0.3\textwidth}}");
            dataFile.WriteLine(@"    \toprule");
            dataFile.WriteLine(@"    \textbf{Remediation} & \textbf{is implementation of: Hygiene Practice} & \textbf{protects against: } \\");


            foreach (IEntity remediation in remediationSortBlock.SortEntities(remediationList))
            {
                dataFile.WriteLine(@"    \midrule");

                var outLine = string.Format(@"    {0}:{1}&",
                        remediation.GetAttribute("number").ToString(), remediation.Name);

                ListCollectionView implementationList = implementationSortBlock.SortEntities(remediation.GetRelationshipTargets("is implementation of"));
                implementationList.MoveCurrentToFirst();
                ListCollectionView protectionList = protectionSortBlock.SortEntities(remediation.GetRelationshipTargets("protects against"));
                protectionList.MoveCurrentToFirst();
                var maxItems = Math.Max(implementationList.Count, protectionList.Count);
                for (var count = 1; count <= maxItems; count++)
                {
                    outLine += Environment.NewLine + @"    ";
                    if (count > 1)
                    {
                        outLine += @"&";
                    }
                    if (implementationList.Count >= count)
                    {
                        IEntity implementation = (IEntity)implementationList.CurrentItem;
                        if (protectionList.Count >= count)
                        {
                            IEntity protection = (IEntity)protectionList.CurrentItem;
                            outLine += string.Format(@"{0}:{1}&{2}:{3}\\",
                                implementation.GetAttribute("number").ToString(), implementation.Name,
                                protection.GetAttribute("number").ToString(), protection.Name);
                            protectionList.MoveCurrentToNext();
                        }
                        else
                        {
                            outLine += string.Format(@"{0}:{1}&\\",
                                implementation.GetAttribute("number").ToString(), implementation.Name);
                        }
                        implementationList.MoveCurrentToNext();
                    }
                    else
                    {
                        IEntity protection = (IEntity)protectionList.CurrentItem;
                        outLine += string.Format(@"&{0}:{1}\\",
                            protection.GetAttribute("number").ToString(), protection.Name);
                        protectionList.MoveCurrentToNext();
                    }
                }
                dataFile.WriteLine(outLine);
            }
            dataFile.WriteLine(@"\bottomrule");
            dataFile.WriteLine(@"\end{tabular}");
            dataFile.WriteLine(@"\label{table:csrm-remediation}");
            dataFile.WriteLine(@"\end{table}");

            dataFile.Close();
        }

        static void ResilientModes(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity resilientModeCategory = cateogoryFolder.GetEntity("CSRM: Resilient Modes");

            IEnumerable<IEntity> resilientModeList = resilientModeCategory.GetRelationshipTargets("categorizes");
            ISortBlock resilientModeSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock lossScenarioSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock componentSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

            // Write Latex Table Definition Header
            dataFile = new System.IO.StreamWriter(outputPath + @"tables/resilientmode.tex", false);
            dataFile.WriteLine(@"% ++++++++++++++++++++");
            dataFile.WriteLine(@"% CSRM Resilient Modes");
            dataFile.WriteLine(@"% ++++++++++++++++++++");

            dataFile.WriteLine(@"\begin{table}[ht]");
            dataFile.WriteLine(@"\scriptsize");
            dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
            dataFile.WriteLine(@"\centering");
            dataFile.WriteLine(@"\caption{" + project.Name + " CSRM Resilient Modes}");
            dataFile.WriteLine(@"\begin{tabular}{l l l}");
            dataFile.WriteLine(@"    \toprule");
            dataFile.WriteLine(@"    \textbf{Resilient Mode} & \textbf{provides reconfiguration for: Loss Scenario} & \textbf{provides alternate operation for: Component} \\");


            foreach (IEntity resilientMode in resilientModeSortBlock.SortEntities(resilientModeList))
            {
                dataFile.WriteLine(@"    \midrule");

                var outLine = string.Format(@"    {0}:{1}&",
                        resilientMode.GetAttribute("number").ToString(), resilientMode.Name);

                ListCollectionView lossScenarioList = 
                    lossScenarioSortBlock.SortEntities(resilientMode.GetRelationshipTargets("provides reconfiguration for"));
                lossScenarioList.MoveCurrentToFirst();
                ListCollectionView componentList = 
                    componentSortBlock.SortEntities(resilientMode.GetRelationshipTargets("provides alternate operation for"));
                componentList.MoveCurrentToFirst();
                var maxItems = Math.Max(lossScenarioList.Count, componentList.Count);
                for (var count = 1; count <= maxItems; count++)
                {
                    outLine += Environment.NewLine + @"    ";
                    if (count > 1)
                    {
                        outLine += @"&";
                    }
                    if (lossScenarioList.Count >= count)
                    {
                        IEntity lossScenario = (IEntity)lossScenarioList.CurrentItem;
                        if (componentList.Count >= count)
                        {
                            IEntity component = (IEntity)componentList.CurrentItem;
                            outLine += string.Format(@"{0}:{1}&{2}:{3}\\",
                                lossScenario.GetAttribute("number").ToString(), lossScenario.Name,
                                component.GetAttribute("number").ToString(), component.Name);
                            componentList.MoveCurrentToNext();
                        }
                        else
                        {
                            outLine += string.Format(@"{0}:{1}&\\",
                                lossScenario.GetAttribute("number").ToString(), lossScenario.Name);
                        }
                        lossScenarioList.MoveCurrentToNext();
                    }
                    else
                    {
                        IEntity component = (IEntity)componentList.CurrentItem;
                        outLine += string.Format(@"&{0}:{1}\\",
                            component.GetAttribute("number").ToString(), component.Name);
                        componentList.MoveCurrentToNext();
                    }
                }
                dataFile.WriteLine(outLine);
            }
            dataFile.WriteLine(@"\bottomrule");
            dataFile.WriteLine(@"\end{tabular}");
            dataFile.WriteLine(@"\label{table:csrm-resilient-modes}");
            dataFile.WriteLine(@"\end{table}");

            dataFile.Close();
        }
        static void SystemDescription( IProject project, string outputPath)
        {
            Architecture(project, outputPath);
            UseCase(project, outputPath);
        }
        static void OperationalRisk( IProject project, string outputPath)
        {
            Losses(project, outputPath);
            Hazards(project, outputPath);
            HazardousActions(project, outputPath);
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
            dataFile.WriteLine(@"    \textbf{Hazard} & \textbf{leads to: Loss} & \textbf{is caused by: Hazardous Action} \\");


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
        static void HazardousActions(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity systemHazardousActionCategory = cateogoryFolder.GetEntity("STPA: Hazardous Action");

            IEnumerable<IEntity> hazardousActionList = systemHazardousActionCategory.GetRelationshipTargets("categorizes");
            ISortBlock haSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock hazardSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock lossScenarioSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);
            ISortBlock controlActionSortBlock = project.GetSortBlock(SortBlockConstants.Numeric);

            // Write Latex Table Definition Header
            dataFile = new System.IO.StreamWriter(outputPath + @"tables/hazardousaction.tex", false);
            dataFile.WriteLine(@"% +++++++++++++++++++++++++++");
            dataFile.WriteLine(@"% STPA Hazardous Action Table");
            dataFile.WriteLine(@"% +++++++++++++++++++++++++++");

            dataFile.WriteLine(@"\begin{table}[ht]");
            dataFile.WriteLine(@"\scriptsize");
            dataFile.WriteLine(@"\renewcommand{\arraystretch}{1.5}");
            dataFile.WriteLine(@"\centering");
            dataFile.WriteLine(@"\caption{" + project.Name + " STPA Hazardous Actions}");
            dataFile.WriteLine(@"\begin{tabular}{p{0.15\textwidth} p{0.2\textwidth} p{0.25\textwidth} p{0.25\textwidth}}");

            dataFile.WriteLine(@"    \toprule");
            dataFile.WriteLine(@"    \textbf{Hazardous Action} & \textbf{leads to: Hazard} & \textbf{is caused by: Loss Scenario} & \textbf{is variation of: Control Action}\\");


            foreach (IEntity hazardousAction in haSortBlock.SortEntities(hazardousActionList))
            {
                dataFile.WriteLine(@"    \midrule");

                var outLine = string.Format(@"    {0}:{1}&",
                        hazardousAction.GetAttribute("number").ToString(), hazardousAction.Name);

                ListCollectionView lossScenarioList = lossScenarioSortBlock.SortEntities(hazardousAction.GetRelationshipTargets("is caused by"));
                lossScenarioList.MoveCurrentToFirst();
                ListCollectionView hazardList = hazardSortBlock.SortEntities(hazardousAction.GetRelationshipTargets("leads to"));
                hazardList.MoveCurrentToFirst();
                ListCollectionView controlActionList = controlActionSortBlock.SortEntities(hazardousAction.GetRelationshipTargets("variation of"));
                controlActionList.MoveCurrentToFirst();

                var maxItems = Math.Max(lossScenarioList.Count, Math.Max(hazardList.Count, controlActionList.Count));
                for (var count = 1; count <= maxItems; count++)
                {
                    outLine += Environment.NewLine + @"    ";
                    if (count > 1)
                    {
                        outLine += @"&";
                    }
                    if (hazardList.Count >= count)
                    {
                        IEntity hazard = (IEntity)hazardList.CurrentItem;
                        if (lossScenarioList.Count >= count)
                        {
                            IEntity lossScenario = (IEntity)lossScenarioList.CurrentItem;
                            if (controlActionList.Count >= count)
                            {
                                IEntity controlAction = (IEntity)controlActionList.CurrentItem;
                                outLine += string.Format(@"{0}:{1}&{2}:{3}&{4}:{5}\\",
                                    hazard.GetAttribute("number").ToString(), hazard.Name,
                                    lossScenario.GetAttribute("number").ToString(), lossScenario.Name,
                                    controlAction.GetAttribute("number").ToString(), controlAction.Name);

                                controlActionList.MoveCurrentToNext();
                            }
                            else
                            {
                                outLine += string.Format(@"{0}:{1}&{2}:{3}&\\",
                                    hazard.GetAttribute("number").ToString(), hazard.Name,
                                    lossScenario.GetAttribute("number").ToString(), lossScenario.Name);
                            }
                            lossScenarioList.MoveCurrentToNext();
                        }
                        else
                        {
                            if (controlActionList.Count >= count)
                            {
                                IEntity controlAction = (IEntity)controlActionList.CurrentItem;
                                outLine += string.Format(@"{0}:{1}&&{2}:{3}\\",
                                    hazard.GetAttribute("number").ToString(), hazard.Name,
                                    controlAction.GetAttribute("number").ToString(), controlAction.Name);

                                controlActionList.MoveCurrentToNext();
                            }
                            else
                            {
                                outLine += string.Format(@"{0}:{1}&&\\",
                                    hazard.GetAttribute("number").ToString(), hazard.Name);
                            }
                        }
                        hazardList.MoveCurrentToNext();
                    }
                    else
                    {
                        if (lossScenarioList.Count >= count)
                        {
                            IEntity lossScenario = (IEntity)lossScenarioList.CurrentItem;
                            if (controlActionList.Count >= count)
                            {
                                IEntity controlAction = (IEntity)controlActionList.CurrentItem;
                                outLine += string.Format(@"&{0}:{1}&{2}:{3}\\",
                                    lossScenario.GetAttribute("number").ToString(), lossScenario.Name,
                                    controlAction.GetAttribute("number").ToString(), controlAction.Name);

                                controlActionList.MoveCurrentToNext();
                            }
                            else
                            {
                                outLine += string.Format(@"&{0}:{1}&\\",
                                    lossScenario.GetAttribute("number").ToString(), lossScenario.Name);
                            }
                            lossScenarioList.MoveCurrentToNext();
                        }
                        else
                        {
                            if (controlActionList.Count >= count)
                            {
                                IEntity controlAction = (IEntity)controlActionList.CurrentItem;
                                outLine += string.Format(@"&&{0}:{1}\\",
                                    controlAction.GetAttribute("number").ToString(), controlAction.Name);

                                controlActionList.MoveCurrentToNext();
                            }
                        }

                        lossScenarioList.MoveCurrentToNext();
                    }
                }
                dataFile.WriteLine(outLine);
            }
            dataFile.WriteLine(@"\bottomrule");
            dataFile.WriteLine(@"\end{tabular}");
            dataFile.WriteLine(@"\label{table:stpa-hazardous-actions}");
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
                OutputDiagram(project, EntityDiagramType.PhysicalBlock, DiagramScale.LineWidth, component, outputPath);
            }
        }
        static void UseCase(IProject project, string outputPath)
        {
            IFolder cateogoryFolder = project.GetFolder("Category");
            IEntity systemUseCaseCategory = cateogoryFolder.GetEntity("CSRM: Use Case");

            IEnumerable<IEntity> usecaseList = systemUseCaseCategory.GetRelationshipTargets("categorizes");

            foreach (IEntity usecase in usecaseList)
            {
                OutputDiagram(project, EntityDiagramType.UseCase, DiagramScale.ThreeQuarter, usecase, outputPath);
            }
        }
        static void OutputDiagram(IProject project, EntityDiagramType diagramType, DiagramScale scale,
            IEntity diagramEntity, string outputPath)
        {
            DiagramExporter diagramExporter = new DiagramExporter(diagramEntity, diagramType, null)
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
