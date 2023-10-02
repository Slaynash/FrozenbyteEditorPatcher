using Mono.Cecil;
using Mono.Cecil.Cil;
using Mono.Cecil.Rocks;
using System.Collections.Generic;
using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.ObjectModel;

namespace FrozenbyteEditorPatcher
{
    public class Program
    {
        public static (string url, string name)[] dockPanelSuiteFiles = new (string, string)[]
        {
            ("https://globalcdn.nuget.org/packages/dockpanelsuite.2.11.0.nupkg"            , "WeifenLuo.WinFormsUI.Docking.dll"),
            ("https://globalcdn.nuget.org/packages/dockpanelsuite.themevs2015.2.11.0.nupkg", "WeifenLuo.WinFormsUI.Docking.ThemeVS2015.dll")
        };

        // Based on 'System.Drawing.KnownColorTable.UpdateSystemColors()' and vs2015dark.vstheme
        // TODO: change '0xFF000000' to the correct colors
        private static (KnownColor key, int value)[] systemColorsOverride = new (KnownColor, int)[]
        {
            (KnownColor.ActiveBorder            , unchecked((int)0xFF3F3F46)),
            (KnownColor.ActiveCaption           , unchecked((int)0xFF2D2D30)),
            (KnownColor.ActiveCaptionText       , unchecked((int)0x99999999)),
            (KnownColor.AppWorkspace            , unchecked((int)0xFF2D2D30)),
            (KnownColor.ButtonFace              , unchecked((int)0xFF3F3F46)),
            (KnownColor.ButtonHighlight         , unchecked((int)0xFF464646)),
            (KnownColor.ButtonShadow            , unchecked((int)0xFF3F3F46)),
            (KnownColor.Control                 , unchecked((int)0xFF000000)),
            (KnownColor.ControlDark             , unchecked((int)0xFF000000)),
            (KnownColor.ControlDarkDark         , unchecked((int)0xFF000000)),
            (KnownColor.ControlLight            , unchecked((int)0xFF000000)),
            (KnownColor.ControlLightLight       , unchecked((int)0xFF000000)),
            (KnownColor.ControlText             , unchecked((int)0xFFF1F1F1)),
            (KnownColor.Desktop                 , unchecked((int)0xFF000000)),
            (KnownColor.GradientActiveCaption   , unchecked((int)0xFF000000)),
            (KnownColor.GradientInactiveCaption , unchecked((int)0xFF000000)),
            (KnownColor.GrayText                , unchecked((int)0xFF999999)),
            (KnownColor.Highlight               , unchecked((int)0xFF3399FF)),
            (KnownColor.HighlightText           , unchecked((int)0xFFFFFFFF)),
            (KnownColor.HotTrack                , unchecked((int)0xFF000000)),
            (KnownColor.InactiveBorder          , unchecked((int)0xFF3F3F46)),
            (KnownColor.InactiveCaption         , unchecked((int)0xFF000000)),
            (KnownColor.InactiveCaptionText     , unchecked((int)0xFF000000)),
            (KnownColor.Info                    , unchecked((int)0xFFFEFCC8)),
            (KnownColor.InfoText                , unchecked((int)0xFF1E1E1E)),
            (KnownColor.Menu                    , unchecked((int)0xFF1B1B1C)),
            (KnownColor.MenuBar                 , unchecked((int)0xFF333337)),
            (KnownColor.MenuHighlight           , unchecked((int)0xFF000000)),
            (KnownColor.MenuText                , unchecked((int)0xFFF1F1F1)),
            (KnownColor.ScrollBar               , unchecked((int)0xFF3E3E42)),
            (KnownColor.Window                  , unchecked((int)0xFF252526)),
            (KnownColor.WindowFrame             , unchecked((int)0xFF2D2D30)),
            (KnownColor.WindowText              , unchecked((int)0xFFF1F1F1))
        };

        private static Color[] logColors = new Color[]
        {
            Color.FromArgb(140, 140, 140), // Debug
            Color.FromArgb(250, 250, 250), // Info
            Color.FromArgb(255, 180,   0), // Warning
            Color.FromArgb(255,  25,  25)  // Error
        };

        private class ColorPatch
        {
            public enum ColorClass { None, Color, SystemColor };
            public Color color;
            public ColorClass colorclass;
            public string colorname;

            public ColorPatch(int r, int g, int b)
            {
                color = Color.FromArgb(r, g, b);
            }

            public ColorPatch(ColorClass colorclass, string colorname)
            {
                this.colorclass = colorclass;
                this.colorname = colorname;
            }
        }


        private static (Func<Instruction, bool>, ColorPatch) patchAllWhiteToWindow =>
            ((inst => inst.OpCode == OpCodes.Call && ((MethodReference)inst.Operand).Name == "get_White"), new ColorPatch(ColorPatch.ColorClass.SystemColor, "Window"));

        // type method (condition, color)
        private static Dictionary<(string namespaze, string name), Dictionary<string, (Func<Instruction, bool> doesMatch, ColorPatch colorpatch)>> individualElementsColorPatchs = new ()
        {
            {
                ("csharpui.window.dockable", "ObjectPropertiesWindow"), new ()
                {
                    { ".ctor"               , patchAllWhiteToWindow },
                    { "endUpdate"           , patchAllWhiteToWindow },
                    { "applySelectedObjects", patchAllWhiteToWindow }
                }
            },
            { ("csharpui.window.dockable.objectexplorer", "SceneExplorer"), new () { { "InitializeComponent", patchAllWhiteToWindow } } },
            {
                ("csharpui.window.dockable.objectexplorer", "ResourceExplorer"), new ()
                {
                    { "InitializeComponent", ((inst => inst.OpCode == OpCodes.Call && ((MethodReference)inst.Operand).Name == "get_Ivory"), new ColorPatch(25, 25, 21)) }
                }
            },
            {
                ("csharpui.window.dockable.objectexplorer", "TypeExplorer"), new ()
                {
                    { "InitializeComponent", ((inst => inst.OpCode == OpCodes.Call && ((MethodReference)inst.Operand).Name == "FromArgb"), new ColorPatch(25, 25, 33)) }
                }
            },
        };


        public static void Main(string[] args)
        {
            string gamePath = @"D:\SteamLibrary\steamapps\common\Shadwen";
            string uiAssemblyPath = Path.Combine(gamePath, "csharpui.dll");
            string dockingAssemblyPath = Path.Combine(gamePath, "WeifenLuo.WinFormsUI.Docking.dll");

            if (File.Exists(uiAssemblyPath + ".patcher_bkp"))
            {
                Console.WriteLine($"Backups already exists, restoring before repatching");
                File.Copy(uiAssemblyPath + ".patcher_bkp", uiAssemblyPath, true);
                File.Copy(dockingAssemblyPath + ".patcher_bkp", dockingAssemblyPath, true);
            }
            else
            {
                Console.WriteLine($"Backing up files");
                File.Copy(uiAssemblyPath, uiAssemblyPath + ".patcher_bkp");
                File.Copy(dockingAssemblyPath, dockingAssemblyPath + ".patcher_bkp");
            }

            List<string> reportLines = new List<string>();

            using (AssemblyDefinition uiAssembly = AssemblyDefinition.ReadAssembly(uiAssemblyPath, new ReaderParameters { ReadWrite = true }))
            {
                DownloadAndPatchDockPanelSuite(uiAssembly, gamePath, reportLines);

                InjectSystemColorsRuntimePatch(uiAssembly);

                InjectDarkWindowBorderPatch(uiAssembly);

                PatchLogMessageColors(uiAssembly);

                PatchTextColorsToBetterOnes(uiAssembly);

                PatchSpecificColorInstructions(uiAssembly);

                uiAssembly.Write();
            }

            Console.WriteLine($"Done!");
        }

        public static void DownloadAndPatchDockPanelSuite(AssemblyDefinition uiAssembly, string gamepath, List<string> reportLines)
        {
            using (HttpClient client = new HttpClient())
            {
                foreach (var file in dockPanelSuiteFiles)
                {
                    Console.WriteLine($"Downloading {file.name}");
                    // Download and extract the nuget package
                    using Task<Stream> streamTask = client.GetStreamAsync(file.url);

                    using ZipArchive zip = new ZipArchive(streamTask.Result);

                    ZipArchiveEntry? entry = zip.GetEntry("lib/net40/" + file.name);
                    if (entry == null)
                        throw new IOException($"file \"lib/net40/{file.name}\" not found in archive.");

                    using FileStream filestream = new FileStream(Path.Combine(gamepath, file.name), FileMode.OpenOrCreate);
                    using Stream zipstream = entry.Open();
                    zipstream.CopyTo(filestream);
                }
            }
            Console.WriteLine($"Patching DockPanelSuite");

            // Change the DockPanel theme to VS2015Dark
            using (AssemblyDefinition dockingAssembly = AssemblyDefinition.ReadAssembly(Path.Combine(gamepath, "WeifenLuo.WinFormsUI.Docking.dll"), new ReaderParameters { ReadWrite = true }))
            {
                using AssemblyDefinition dockingThemeAssembly = AssemblyDefinition.ReadAssembly(Path.Combine(gamepath, "WeifenLuo.WinFormsUI.Docking.ThemeVS2015.dll"));

                MethodReference vs2015ThemeConstructorRef = dockingThemeAssembly.MainModule
                    .Types.Single(typeref => typeref.Namespace == "WeifenLuo.WinFormsUI.Docking" && typeref.Name == "VS2015DarkTheme")
                    .GetConstructors().Single(c => !c.IsStatic);

                // Add an assembly reference to ThemeVS2015
                //dockingAssembly.MainModule.AssemblyReferences.Add(dockingThemeAssembly.Name);
                vs2015ThemeConstructorRef = dockingAssembly.MainModule.ImportReference(vs2015ThemeConstructorRef);

                dockingAssembly.MainModule
                    .Types.Single(typeref => typeref.Namespace == "WeifenLuo.WinFormsUI.Docking" && typeref.Name == "DockPanel")
                    .GetConstructors().Single(c => !c.IsStatic)
                    .Body.Instructions.First(inst => inst.OpCode == OpCodes.Newobj && ((MethodReference)inst.Operand).DeclaringType.Name == "VS2005Theme")
                    .Operand = vs2015ThemeConstructorRef;

                dockingAssembly.Write();
            }

            // Update the reference of DockingPanelSuite from 2.8 to 2.11
            Console.WriteLine($"Patching DockPanelSuite version info of csharpui");
            uiAssembly.MainModule.AssemblyReferences.Single(aref => aref.Name == "WeifenLuo.WinFormsUI.Docking").Version = new Version(2, 11, 0);


            // Remove the set_Skin call from 'csharpui.window.main.CSharpUIMDI.InitializeComponent()'
            // This is temporary, the missing customization should be re-patched in
            MethodDefinition mdiInitalizeComponentMethod = uiAssembly.MainModule
                .Types.Single(t => t.Namespace == "csharpui.window.main" && t.Name == "CSharpUIMDI")
                .Methods.Single(m => m.Name == "InitializeComponent");

            var mdiInitalizeComponentInstructions = mdiInitalizeComponentMethod.Body.Instructions;
            int setSkinFirstInstructionOffset = mdiInitalizeComponentInstructions.IndexOf(
                mdiInitalizeComponentInstructions.Single(inst =>
                    inst               .OpCode == OpCodes.Ldarg_0  &&
                    inst.Next          .OpCode == OpCodes.Ldfld    && ((FieldReference) inst.Next          .Operand).Name == "mainDockPanel" &&
                    inst.Next.Next     .OpCode == OpCodes.Ldloc_1  &&
                    inst.Next.Next.Next.OpCode == OpCodes.Callvirt && ((MethodReference)inst.Next.Next.Next.Operand).Name == "set_Skin"));

            mdiInitalizeComponentMethod.Body.Instructions.RemoveAt(setSkinFirstInstructionOffset);
            mdiInitalizeComponentMethod.Body.Instructions.RemoveAt(setSkinFirstInstructionOffset);
            mdiInitalizeComponentMethod.Body.Instructions.RemoveAt(setSkinFirstInstructionOffset);
            mdiInitalizeComponentMethod.Body.Instructions.RemoveAt(setSkinFirstInstructionOffset);
        }

        public static void InjectSystemColorsRuntimePatch(AssemblyDefinition uiAssembly)
        {
            // Fetch the required references

            Console.WriteLine($"Resolving references to inject colors patch");
            ModuleDefinition systemDrawingModule = uiAssembly.MainModule.AssemblyResolver.Resolve(uiAssembly.MainModule.AssemblyReferences.Single(aref => aref.Name == "System.Drawing")).MainModule;
            TypeDefinition colorTypeDef = systemDrawingModule.Types.Single(typeref => typeref.Namespace == "System.Drawing" && typeref.Name == "Color");
            TypeReference colorType = uiAssembly.MainModule.ImportReference(colorTypeDef);
            MethodReference toArgbMethod = uiAssembly.MainModule.ImportReference(colorTypeDef.Methods.Single(m => m.Name == "ToArgb"));
            TypeDefinition systemcolorsType = systemDrawingModule.Types.Single(typeref => typeref.Namespace == "System.Drawing" && typeref.Name == "SystemColors");
            MethodReference getWindowMethod = uiAssembly.MainModule.ImportReference(systemcolorsType.Methods.Single(m => m.Name == "get_Window"));

            ModuleDefinition mscorlibModule = uiAssembly.MainModule.AssemblyResolver.Resolve(uiAssembly.MainModule.AssemblyReferences.First(aref => aref.Name == "mscorlib")).MainModule;
            TypeReference int32Type = uiAssembly.MainModule.ImportReference(mscorlibModule.Types.Single(t => t.Namespace == "System" && t.Name == "Int32"));
            TypeReference int32ArrayType = uiAssembly.MainModule.ImportReference(new ArrayType(int32Type));
            TypeDefinition typeType = mscorlibModule.Types.Single(t => t.Namespace == "System" && t.Name == "Type");
            MethodReference getTypeFromHandleMethod = uiAssembly.MainModule.ImportReference(typeType.Methods.Single(m => m.Name == "GetTypeFromHandle"));
            MethodReference getassemblyMethod = uiAssembly.MainModule.ImportReference(typeType.Methods.Single(m => m.Name == "get_Assembly"));
            MethodReference getFieldMethod = uiAssembly.MainModule.ImportReference(typeType.Methods.Single(m => m.Name == "GetField" && m.Parameters.Count == 2));

            TypeDefinition assemblyType = mscorlibModule.Types.Single(typeref => typeref.Namespace == "System.Reflection" && typeref.Name == "Assembly");
            MethodReference getTypeMethod = uiAssembly.MainModule.ImportReference(assemblyType.Methods.Single(m => m.Name == "GetType" && m.Parameters.Count == 1));

            TypeDefinition fieldinfoType = mscorlibModule.Types.Single(typeref => typeref.Namespace == "System.Reflection" && typeref.Name == "FieldInfo");
            MethodReference getValueMethod = uiAssembly.MainModule.ImportReference(fieldinfoType.Methods.Single(m => m.Name == "GetValue" && m.Parameters.Count == 1));


            // The type System.Drawing.KnownColorTable and its field colorTable are both non-public. We get them by reflection.
            // We make sure everything already ran by fetching a random color.
            //   SystemColors.Window.ToArgb()
            //   typeof(Color).Assembly.GetType("System.Drawing.KnownColorTable").GetField("colorTable", (BindingFlags)(-1)).GetValue(null) as int[]

            MethodDefinition createMDIWindowMethodRef = uiAssembly.MainModule
                .Types.Single(typeref => typeref.Namespace == "csharpui" && typeref.Name == "CSharpUI")
                .Methods.Single(methodref => methodref.Name == "createMDIWindow");

            Console.WriteLine($"Injecting color patch");

            {
                ILProcessor ilprocessor = createMDIWindowMethodRef.Body.GetILProcessor();
                Instruction firstBranchInstruction = createMDIWindowMethodRef.Body.Instructions.First(inst => inst.OpCode == OpCodes.Brtrue_S);
                firstBranchInstruction.OpCode = OpCodes.Brtrue; // We're too far for a short branch
                Instruction firstPostBranchInstruction = firstBranchInstruction.Next;

                VariableDefinition colorVariable = new VariableDefinition(colorType);
                createMDIWindowMethodRef.Body.Variables.Add(colorVariable);
                // SystemColors.Window.ToArgb()
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Call, getWindowMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Stloc_S, colorVariable));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldloca_S, colorVariable));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Call, toArgbMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Pop));
                // loc2 = typeof(Color).Assembly.GetType("System.Drawing.KnownColorTable").GetField("colorTable", (BindingFlags)(-1)).GetValue(null) as int[]
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldtoken, colorType));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Call, getTypeFromHandleMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Callvirt, getassemblyMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldstr, "System.Drawing.KnownColorTable"));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Callvirt, getTypeMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldstr, "colorTable"));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldc_I4_M1));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Callvirt, getFieldMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldnull));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Callvirt, getValueMethod));
                ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Isinst, int32ArrayType));
                // foreach pair in systemColorsOverride: loc2[pair.key] = pair.value
                for (int i = 0; i < systemColorsOverride.Length; ++i)
                {
                    var pair = systemColorsOverride[i];
                    if (i != systemColorsOverride.Length - 1)
                        ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Dup));
                    ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldc_I4, (int)pair.key));
                    ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Ldc_I4, pair.value));
                    ilprocessor.InsertBefore(firstPostBranchInstruction, ilprocessor.Create(OpCodes.Stelem_I4));
                }
            }


        }
        public static void InjectDarkWindowBorderPatch(AssemblyDefinition uiAssembly)
        {
            Console.WriteLine($"Resolving references to inject dark window patch");
            ModuleDefinition mscorlibModule = uiAssembly.MainModule.AssemblyResolver.Resolve(uiAssembly.MainModule.AssemblyReferences.First(aref => aref.Name == "mscorlib")).MainModule;
            TypeReference int32Type = uiAssembly.MainModule.ImportReference(mscorlibModule.Types.Single(t => t.Namespace == "System" && t.Name == "Int32"));
            TypeReference int32ByrefType = uiAssembly.MainModule.ImportReference(new ByReferenceType(int32Type));
            TypeReference intType = uiAssembly.MainModule.ImportReference(mscorlibModule.Types.Single(t => t.Namespace == "System" && t.Name == "IntPtr"));


            ModuleDefinition winformModule = uiAssembly.MainModule.AssemblyResolver.Resolve(uiAssembly.MainModule.AssemblyReferences.First(aref => aref.Name == "System.Windows.Forms")).MainModule;
            TypeDefinition formType = winformModule.Types.Single(typeref => typeref.Namespace == "System.Windows.Forms" && typeref.Name == "Control");
            MethodReference getHandleMethod = uiAssembly.MainModule.ImportReference(formType.Methods.Single(m => m.Name == "get_Handle"));

            MethodDefinition mdiConstructor = uiAssembly.MainModule
                .Types.Single(t => t.Namespace == "csharpui.window.main" && t.Name == "CSharpUIMDI")
                .GetConstructors().Single(c => !c.IsStatic);


            Console.WriteLine($"Injecting DwmSetWindowAttribute");

            ModuleReference dwmapiModule = new ModuleReference("dwmapi.dll");
            uiAssembly.MainModule.ModuleReferences.Add(dwmapiModule);

            MethodDefinition dwmSetWindowAttributeMethod = new MethodDefinition("DwmSetWindowAttribute", MethodAttributes.Static | MethodAttributes.HideBySig, int32Type);
            dwmSetWindowAttributeMethod.Parameters.Add(new ParameterDefinition("hwnd",      ParameterAttributes.None, intType));
            dwmSetWindowAttributeMethod.Parameters.Add(new ParameterDefinition("attr",      ParameterAttributes.None, int32Type));
            dwmSetWindowAttributeMethod.Parameters.Add(new ParameterDefinition("attrValue", ParameterAttributes.None, int32ByrefType));
            dwmSetWindowAttributeMethod.Parameters.Add(new ParameterDefinition("attrSize",  ParameterAttributes.None, int32Type));
            dwmSetWindowAttributeMethod.ImplAttributes = MethodImplAttributes.PreserveSig;
            dwmSetWindowAttributeMethod.PInvokeInfo = new PInvokeInfo(PInvokeAttributes.CallConvWinapi, "DwmSetWindowAttribute", dwmapiModule);
            mdiConstructor.DeclaringType.Methods.Add(dwmSetWindowAttributeMethod);

            Console.WriteLine($"Injecting dark window patch");

            {
                ILProcessor ilprocessor = mdiConstructor.Body.GetILProcessor();
                Instruction callInitializeComponentPostInstruction = mdiConstructor.Body
                    .Instructions.First(inst => inst.OpCode == OpCodes.Call && ((MethodReference)inst.Operand).Name == "InitializeComponent")
                    .Next;

                VariableDefinition attrValueVariable = new VariableDefinition(int32Type);
                mdiConstructor.Body.Variables.Add(attrValueVariable);

                // int attrValue = 1
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Ldc_I4_1));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Stloc_S, attrValueVariable));
                // DwmSetWindowAttribute(Handle, 20, ref attrValue, 4);
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Ldarg_0));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Call, getHandleMethod));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Ldc_I4_S, (sbyte)20));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Ldloca_S, attrValueVariable));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Ldc_I4_4));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Call, dwmSetWindowAttributeMethod));
                ilprocessor.InsertBefore(callInitializeComponentPostInstruction, ilprocessor.Create(OpCodes.Pop));
            }
        }

        public static void PatchLogMessageColors(AssemblyDefinition uiAssembly)
        {
            MethodDefinition logMessageImplMethod = uiAssembly.MainModule
                .Types.Single(t => t.Namespace == "csharpui.window.dockable" && t.Name == "OutputWindow")
                .Methods.Single(m => m.Name == "logMessageQueuePurgeImpl");

            Instruction textBaseStringInstruction = logMessageImplMethod.Body.Instructions.Single(inst => inst.OpCode == OpCodes.Ldstr && ((string)inst.Operand).Contains("colortbl"));
            textBaseStringInstruction.Operand = "{" +
                @"\rtf\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Courier New;}}" +
                @"{\colortbl;" +
                    string.Join("",logColors.Select(c => $@"\red{c.R}\green{c.G}\blue{c.B};")) +
                @"}\fs16";
        }

        public static void PatchTextColorsToBetterOnes(AssemblyDefinition uiAssembly)
        {
            ModuleDefinition systemDrawingModule = uiAssembly.MainModule
                .AssemblyResolver.Resolve(
                    uiAssembly.MainModule.AssemblyReferences.Single(aref => aref.Name == "System.Drawing"))
                .MainModule;
            TypeDefinition colorTypeDef = systemDrawingModule.Types.Single(typeref => typeref.Namespace == "System.Drawing" && typeref.Name == "Color");
            TypeReference colorTypeRef = uiAssembly.MainModule.ImportReference(systemDrawingModule.Types.Single(typeref => typeref.Namespace == "System.Drawing" && typeref.Name == "Color"));

            MethodReference getCyanMethod = uiAssembly.MainModule.ImportReference(colorTypeDef.Methods.Single(m => m.Name == "get_Cyan"));
            MethodReference getGreenYellowMethod = uiAssembly.MainModule.ImportReference(colorTypeDef.Methods.Single(m => m.Name == "get_GreenYellow"));

            foreach (TypeDefinition type in uiAssembly.MainModule.Types)
            {
                foreach (MethodDefinition methodDefinition in type.Methods)
                {
                    if (!methodDefinition.HasBody)
                        continue;

                    foreach (Instruction inst in methodDefinition.Body.Instructions)
                    {
                        if (inst.OpCode == OpCodes.Call)
                        {
                            MethodReference mref = inst.Operand as MethodReference;
                            if (mref.DeclaringType.Name != "Color")
                                continue;

                            switch (mref.Name)
                            {
                                case "get_Blue": inst.Operand = getCyanMethod; break;
                                case "get_Green": inst.Operand = getGreenYellowMethod; break;
                            }
                        }
                    }
                }
            }

        }

        public static void PatchSpecificColorInstructions(AssemblyDefinition uiAssembly)
        {
            ModuleDefinition systemDrawingModule = uiAssembly.MainModule
                .AssemblyResolver.Resolve(
                    uiAssembly.MainModule.AssemblyReferences.Single(aref => aref.Name == "System.Drawing"))
                .MainModule;

            TypeDefinition colorTypeDef = systemDrawingModule.Types.Single(typeref => typeref.Namespace == "System.Drawing" && typeref.Name == "Color");
            TypeDefinition systemcolorsType = systemDrawingModule.Types.Single(typeref => typeref.Namespace == "System.Drawing" && typeref.Name == "SystemColors");
            MethodReference fromArgbMethod = uiAssembly.MainModule.ImportReference(
                colorTypeDef.Methods.Single(m => m.Name == "FromArgb" && m.Parameters.Count == 3));



            Dictionary<ColorPatch.ColorClass, Dictionary<string, MethodReference>> colorMethodReferences = new()
            {
                { ColorPatch.ColorClass.Color      , new () },
                { ColorPatch.ColorClass.SystemColor, new () }
            };

            // Populate colorMethodReferences
            foreach (var individualElementsColorPatchType in individualElementsColorPatchs)
            {
                foreach (var individualElementsColorPatchMethods in individualElementsColorPatchType.Value)
                {
                    ColorPatch colorpatch = individualElementsColorPatchMethods.Value.colorpatch;
                    if (colorpatch.colorclass == ColorPatch.ColorClass.None || colorMethodReferences[colorpatch.colorclass].ContainsKey(colorpatch.colorname))
                        continue;


                    TypeDefinition typedef = colorpatch.colorclass == ColorPatch.ColorClass.Color
                        ? colorTypeDef
                        : systemcolorsType;

                    colorMethodReferences[colorpatch.colorclass][colorpatch.colorname] = uiAssembly.MainModule.ImportReference(typedef.Methods.Single(m => m.Name == "get_" + colorpatch.colorname));
                }
            }

            foreach (var individualElementsColorPatchType in individualElementsColorPatchs)
            {
                var typefullname = individualElementsColorPatchType.Key;
                TypeDefinition typedef = uiAssembly.MainModule.Types.Single(t => t.Namespace == typefullname.namespaze && t.Name == typefullname.name);
                foreach (var individualElementsColorPatchMethods in individualElementsColorPatchType.Value)
                {
                    foreach (MethodDefinition methoddef in typedef.Methods)
                    {
                        if (individualElementsColorPatchMethods.Key != methoddef.Name)
                            continue;

                        var instructions = methoddef.Body.Instructions;
                        ILProcessor ilprocessor = methoddef.Body.GetILProcessor();
                        for (int i = 0; i < instructions.Count; i++)
                        {
                            Instruction inst = instructions[i];
                            var (doesMatch, colorpatch) = individualElementsColorPatchMethods.Value;
                            if (!doesMatch(inst))
                                continue;

                            if (colorpatch.colorclass == ColorPatch.ColorClass.None)
                            {
                                if (((MethodReference)inst.Operand).Name == "FromArgb")
                                {
                                    if (
                                        inst.Previous.Previous.Previous.OpCode == OpCodes.Ldc_I4 &&
                                        inst.Previous.Previous         .OpCode == OpCodes.Ldc_I4 &&
                                        inst.Previous                  .OpCode == OpCodes.Ldc_I4
                                    )
                                    {
                                        inst.Previous.Previous.Previous.Operand = (int)colorpatch.color.R;
                                        inst.Previous.Previous         .Operand = (int)colorpatch.color.G;
                                        inst.Previous                  .Operand = (int)colorpatch.color.B;
                                    }
                                    else
                                        throw new NotImplementedException("Patching a FromArgb requires the 3 previous instructions to all be ldc.i4");
                                }
                                else
                                {
                                    ilprocessor.InsertBefore(inst, ilprocessor.Create(OpCodes.Ldc_I4, (int)colorpatch.color.R));
                                    ilprocessor.InsertBefore(inst, ilprocessor.Create(OpCodes.Ldc_I4, (int)colorpatch.color.G));
                                    ilprocessor.InsertBefore(inst, ilprocessor.Create(OpCodes.Ldc_I4, (int)colorpatch.color.B));
                                    ilprocessor.InsertBefore(inst, ilprocessor.Create(OpCodes.Call, fromArgbMethod));
                                    ilprocessor.Remove(inst);
                                    i += 3;
                                }
                            }
                            else
                            {
                                if (((MethodReference)inst.Operand).Name == "FromArgb")
                                {
                                    if (
                                        inst.Previous.Previous.Previous.OpCode == OpCodes.Ldc_I4 &&
                                        inst.Previous.Previous.OpCode == OpCodes.Ldc_I4 &&
                                        inst.Previous.OpCode == OpCodes.Ldc_I4
                                    )
                                    {
                                        ilprocessor.InsertAfter(inst, ilprocessor.Create(OpCodes.Call, colorMethodReferences[colorpatch.colorclass][colorpatch.colorname]));
                                        ilprocessor.Remove(inst.Previous.Previous.Previous);
                                        ilprocessor.Remove(inst.Previous.Previous);
                                        ilprocessor.Remove(inst.Previous);
                                        ilprocessor.Remove(inst);
                                        i -= 3;
                                    }
                                    else
                                        throw new NotImplementedException("Patching a FromArgb requires the 3 previous instructions to all be ldc.i4");
                                }
                                else
                                {
                                    inst.Operand = colorMethodReferences[colorpatch.colorclass][colorpatch.colorname];
                                }
                            }
                        }
                    }
                }
            }

        }

    }
}