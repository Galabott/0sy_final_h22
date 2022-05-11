using ExcelToExcel.Models;
using ExcelToExcel.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Xunit;

namespace ExcelToExcel.Tests
{
    public class MainViewModelTests
    {
        
        string excelFilesPath;

        MainViewModel vm;

        public MainViewModelTests()
        {
            vm = new MainViewModel();

            Uri codeBaseUrl = new Uri(Assembly.GetExecutingAssembly().Location);
            var codeBasePath = Uri.UnescapeDataString(codeBaseUrl.AbsolutePath);
            var dirPath = Path.GetDirectoryName(codeBasePath);

            /// Va chercher le dossier Data à partir du dossier de compilation
            /// Adapter selon votre réalité
            excelFilesPath = Path.Combine(dirPath, @"..\..\..\..\..\data");
        }

        private void resetData()
        {
            vm.InputFilename = "";
        }

        #region Tests à compléter

        /// <summary>
        /// Le fichier d'entrée est vide. La propriété Message devrait être vide.
        /// </summary>
        [Fact]
        public void InputFile_IsEmpty_Message_ShouldBe_Empty()
        {
            /// xTODO : Q01a. Compléter le test
            /// xTODO : Q01b. Ne pas briser la batterie de tests après ce tests
            vm.LoadContentCommand.CanExecute("");

            var actual = vm.Message;
            var expected = "";

            Assert.Equal(expected, actual);
        }

        // xTODO : Q02 : Créer le test CanExecuteSaveCommand_FileNotLoaded_ShouldReturn_False
        [Theory]
        [MemberData(nameof(ExcelFilesTestData))]
        public void CanExecuteSaveCommand_FileNotLoaded_ShouldReturn_False(string fn)
        {
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            var actual = vm.SaveCommand.CanExecute("");

            /// Assert
            Assert.False(actual);
        }

        // xTODO : Q03 : Créer le test CanExecuteSaveCommand_OutputFileInvalid_ShouldReturn_False
        [Theory]
        [MemberData(nameof(WrongOutputFileNames))]
        public void CanExecuteSaveCommand_OutputFileInvalid_ShouldReturn_False(string output)
        {
            var filename = Path.Combine(excelFilesPath, "liste_especes.xlsx");
            vm.InputFilename = filename;
            /// Act
            vm.LoadContentCommand.Execute(null);
            vm.OutputFilename = output;

            var actual = vm.SaveCommand.CanExecute("");

            /// Assert
            Assert.False(actual);
        }

        // xTODO : Q04 : Créer le test CanExecuteSaveCommand_OutputFileValid_ShouldReturn_True(string filename)
        [Theory]
        [MemberData(nameof(GoodOutputFileNames))]
        public void CanExecuteSaveCommand_OutputFileValid_ShouldReturn_True(string output)
        {
            var filename = Path.Combine(excelFilesPath, "liste_especes.xlsx");
            vm.InputFilename = filename;
            /// Act
            vm.LoadContentCommand.Execute(null);
            vm.OutputFilename = output;

            var actual = vm.SaveCommand.CanExecute("");

            /// Assert
            Assert.False(actual);
        }

        #endregion

        #region Tests corrigés

        [Theory]
        [MemberData(nameof(ExistingFilesTestData))]
        public void LoadContentCommand_CanExecute_FileExists_ShouldReturn_true(string fn)
        {            
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            var actual = vm.LoadContentCommand.CanExecute("");

            /// Assert
            Assert.True(actual);
        }

        [Theory]
        [MemberData(nameof(NonExistentFilesTestData))]
        public void LoadContentCommand_CanExecute_FileNotExists_ShouldReturn_false(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            var actual = vm.LoadContentCommand.CanExecute("");

            /// Assert
            Assert.False(actual);
        }

        [Theory]
        [MemberData(nameof(NonExistentFilesTestData))]
        public void LoadContentCommand_CanExecute_FileNotExists_Message_ShouldBe_FileNonExistent(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;
            var expected = "Fichier inexistant!";

            /// Act
            vm.LoadContentCommand.CanExecute("");
            var actual = vm.Message;

            /// Assert
            Assert.Equal(expected, actual);
        }

        [Theory]
        [MemberData(nameof(BadFileTypesTestData))]
        public void LoadContentCommand_FileWrongFormat_Message_ShouldBe_BadFormat(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;
            var expected = "Mauvais format de fichier!";

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.Message;

            /// Assert
            Assert.Equal(expected, actual);
        }

        [Theory]
        [MemberData(nameof(BadExcelFilesTestData))]
        public void LoadContentCommand_ExcelWrongFormat_Message_ShouldBe_BadFormat(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;
            var expected = "Mauvais format de fichier!";

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.Message;

            /// Assert
            Assert.Equal(expected, actual);
        }

        [Theory]
        [MemberData(nameof(GoodExcelFileTestData))]
        public void LoadContentCommand_GoodExcel_Should_HaveContent(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.FileContent;

            /// Assert
            Assert.NotEqual("", actual);
        }


        [Theory]
        [MemberData(nameof(GoodExcelFileTestData))]
        public void LoadContentCommand_ReadOnlyExcel_ShouldDisplay_Message(string fn)
        {
            /// Arrange
            var filename = Path.Combine(excelFilesPath, fn);
            vm.InputFilename = filename;

            /// Act
            vm.LoadContentCommand.Execute(null);
            var actual = vm.FileContent;

            /// Assert
            Assert.NotEqual("", actual);
        }

        #endregion


        #region TEST DATA
        public static IEnumerable<object[]> ExistingFilesTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
            new object[] {"invalide_fichier_type.txt"},
        };

        public static IEnumerable<object[]> ExcelFilesTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"liste_especes_multifeuilles.xlsx"},
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
        };

        public static IEnumerable<object[]> BadExcelFilesTestData = new List<object[]>
        {
            new object[] {"Contenu_nom de peuplement.xlsx"},
            new object[] {"faune_aquatique_v21.xlsx"},
            new object[] {"faune_aquatique_v21_segment.xlsx"},
            new object[] {"Tableau_Export_v1.xlsx"},
        };

        public static IEnumerable<object[]> NonExistentFilesTestData = new List<object[]>
        {
            new object[] {"a.xlsx"},
            new object[] {"test.txt"},
            new object[] {"fasfdsfa.xlsx"},
            new object[] {"loooolollol.xlsx"},
        };

        public static IEnumerable<object[]> WrongOutputFileNames = new List<object[]>
        {
            new object[] {"!841!.xlsx"},
            new object[] {"!%#!&*!.txt"},
            new object[] {")(*&^!.csv"},
            new object[] {"loooolollol.txt"},
        };
        public static IEnumerable<object[]> GoodOutputFileNames = new List<object[]>
        {
            new object[] {"a.xlsx"},
            new object[] {"lol.txt"},
            new object[] {"yeet.csv"},
            new object[] {"oyeee.txt"},
        };

        public static IEnumerable<object[]> BadFileTypesTestData = new List<object[]>
        {
            new object[] {"invalide_fichier_type.txt"},
        };

        public static IEnumerable<object[]> GoodExcelFileTestData = new List<object[]>
        {
            new object[] {"liste_especes.xlsx"},
            new object[] {"liste_especes_multifeuilles.xlsx"},           
        };


        #endregion
    }
}
