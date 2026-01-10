using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace Quoc_MEP.Tests.SampleTests
{
    /// <summary>
    /// Các test cases cơ bản KHÔNG cần Revit API
    /// Dùng để demo NUnit testing framework
    /// </summary>
    [TestFixture]
    public class BasicTests
    {
        [SetUp]
        public void Setup()
        {
            // Khởi tạo trước mỗi test
        }

        [Test]
        [Category("Basic")]
        [Description("Test cộng hai số")]
        public void TestAddition()
        {
            // Arrange
            int a = 5;
            int b = 3;
            int expected = 8;

            // Act
            int result = a + b;

            // Assert
            Assert.AreEqual(expected, result);
        }

        [Test]
        [Category("Basic")]
        [Description("Test string concatenation")]
        public void TestStringConcat()
        {
            // Arrange
            string first = "Hello";
            string second = "World";
            string expected = "HelloWorld";

            // Act
            string result = first + second;

            // Assert
            Assert.AreEqual(expected, result);
        }

        [Test]
        [Category("Math")]
        [Description("Test square root calculation")]
        public void TestSquareRoot()
        {
            // Arrange
            double input = 16.0;
            double expected = 4.0;

            // Act
            double result = Math.Sqrt(input);

            // Assert
            Assert.AreEqual(expected, result, 0.0001);
        }

        [Test]
        [Category("Math")]
        [Description("Test pythagorean theorem")]
        public void TestPythagorean()
        {
            // Arrange - tam giác vuông 3-4-5
            double a = 3.0;
            double b = 4.0;
            double expectedHypotenuse = 5.0;

            // Act
            double hypotenuse = Math.Sqrt(a * a + b * b);

            // Assert
            Assert.AreEqual(expectedHypotenuse, hypotenuse, 0.0001, 
                "Hypotenuse của tam giác 3-4-5 phải là 5");
        }

        [Test]
        [Category("Conversion")]
        [Description("Test feet to millimeters conversion")]
        public void TestFeetToMillimeters()
        {
            // Arrange
            double feet = 1.0;
            double mmPerFoot = 304.8;
            double expectedMm = 304.8;

            // Act
            double mm = feet * mmPerFoot;

            // Assert
            Assert.AreEqual(expectedMm, mm, 0.0001);
        }

        [Test]
        [Category("Conversion")]
        [Description("Test millimeters to feet conversion")]
        public void TestMillimetersToFeet()
        {
            // Arrange
            double mm = 304.8;
            double mmPerFoot = 304.8;
            double expectedFeet = 1.0;

            // Act
            double feet = mm / mmPerFoot;

            // Assert
            Assert.AreEqual(expectedFeet, feet, 0.0001);
        }

        [Test]
        [Category("Collection")]
        [Description("Test list operations")]
        public void TestListOperations()
        {
            // Arrange
            List<int> numbers = new List<int> { 1, 2, 3, 4, 5 };

            // Assert
            Assert.AreEqual(5, numbers.Count, "List should have 5 elements");
            Assert.Contains(3, numbers, "List should contain 3");
            Assert.IsTrue(numbers.Contains(5), "List should contain 5");
        }

        [Test]
        [Category("Validation")]
        [Description("Test null checks")]
        public void TestNullValidation()
        {
            // Arrange
            string nullString = null;
            string validString = "Not null";

            // Assert
            Assert.IsNull(nullString, "String should be null");
            Assert.IsNotNull(validString, "String should not be null");
        }

        [Test]
        [Category("Exception")]
        [Description("Test exception handling")]
        public void TestDivisionByZero()
        {
            // Assert
            Assert.Throws<DivideByZeroException>(() =>
            {
                int a = 10;
                int b = 0;
                int result = a / b;
            });
        }

        [TearDown]
        public void TearDown()
        {
            // Dọn dẹp sau mỗi test
        }
    }

    /// <summary>
    /// Test helper class cho geometry calculations
    /// </summary>
    public class GeometryHelper
    {
        public static double CalculateDistance2D(double x1, double y1, double x2, double y2)
        {
            double dx = x2 - x1;
            double dy = y2 - y1;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        public static double CalculateDistance3D(double x1, double y1, double z1, 
                                                  double x2, double y2, double z2)
        {
            double dx = x2 - x1;
            double dy = y2 - y1;
            double dz = z2 - z1;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }
    }

    /// <summary>
    /// Tests cho GeometryHelper
    /// </summary>
    [TestFixture]
    public class GeometryHelperTests
    {
        [Test]
        [Category("Geometry")]
        [Description("Test 2D distance calculation")]
        public void TestDistance2D()
        {
            // Arrange
            double x1 = 0, y1 = 0;
            double x2 = 3, y2 = 4;
            double expected = 5.0;

            // Act
            double result = GeometryHelper.CalculateDistance2D(x1, y1, x2, y2);

            // Assert
            Assert.AreEqual(expected, result, 0.0001);
        }

        [Test]
        [Category("Geometry")]
        [Description("Test 3D distance calculation")]
        public void TestDistance3D()
        {
            // Arrange
            double x1 = 0, y1 = 0, z1 = 0;
            double x2 = 2, y2 = 2, z2 = 1;
            double expected = 3.0;

            // Act
            double result = GeometryHelper.CalculateDistance3D(x1, y1, z1, x2, y2, z2);

            // Assert
            Assert.AreEqual(expected, result, 0.0001);
        }
    }
}
