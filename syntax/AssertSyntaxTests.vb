Option Explicit On
Imports NUnit.Framework

Namespace NUnit.Samples

    ' This test fixture attempts to exercise all the syntactic
    ' variations of Assert without getting into failures, errors 
    ' or corner cases. Thus, some of the tests may be duplicated 
    ' in other fixtures.
    ' 
    ' Each test performs the same operations using the classic
    ' syntax (if available) and the new syntax.
    ' 
    ' This Fixture will eventually be duplicated in other
    ' supported languages. 

    <TestFixture()>
    Public Class AssertSyntaxTests

#Region "Simple Constraint Tests"
        <Test()>
        Public Sub IsNull()
            Dim nada As Object = Nothing

            ' Classic syntax
            Assert.IsNull(nada)

            ' Helper syntax
            Assert.That(nada, [Is].Null)
        End Sub


        <Test()>
        Public Sub IsNotNull()
            ' Classic syntax
            Assert.IsNotNull(42)

            ' Helper syntax
            Assert.That(42, [Is].Not.Null)
        End Sub

        <Test()>
        Public Sub IsTrue()
            ' Classic syntax
            Assert.IsTrue(2 + 2 = 4)

            ' Helper syntax
            Assert.That(2 + 2 = 4, [Is].True)
            Assert.That(2 + 2 = 4)
        End Sub

        <Test()>
        Public Sub IsFalse()
            ' Classic syntax
            Assert.IsFalse(2 + 2 = 5)

            ' Helper syntax
            Assert.That(2 + 2 = 5, [Is].False)
        End Sub

        <Test()>
        Public Sub IsNaN()
            Dim d As Double = Double.NaN
            Dim f As Single = Single.NaN

            ' Classic syntax
            Assert.IsNaN(d)
            Assert.IsNaN(f)

            ' Helper syntax
            Assert.That(d, [Is].NaN)
            Assert.That(f, [Is].NaN)
        End Sub

        <Test()>
        Public Sub EmptyStringTests()
            ' Classic syntax
            Assert.IsEmpty("")
            Assert.IsNotEmpty("Hello!")

            ' Helper syntax
            Assert.That("", [Is].Empty)
            Assert.That("Hello!", [Is].Not.Empty)
        End Sub

        <Test()>
        Public Sub EmptyCollectionTests()

            Dim boolArray As Boolean() = New Boolean() {}
            Dim nonEmpty As Integer() = New Integer() {1, 2, 3}

            ' Classic syntax
            Assert.IsEmpty(boolArray)
            Assert.IsNotEmpty(nonEmpty)

            ' Helper syntax
            Assert.That(boolArray, [Is].Empty)
            Assert.That(nonEmpty, [Is].Not.Empty)
        End Sub
#End Region

#Region "TypeConstraint Tests"
        <Test()>
        Public Sub ExactTypeTests()
            ' Classic syntax workarounds
            Assert.AreEqual(GetType(String), "Hello".GetType())
            Assert.AreEqual("System.String", "Hello".GetType().FullName)
            Assert.AreNotEqual(GetType(Integer), "Hello".GetType())
            Assert.AreNotEqual("System.Int32", "Hello".GetType().FullName)

            ' Helper syntax
            Assert.That("Hello", [Is].TypeOf(GetType(String)))
            Assert.That("Hello", [Is].Not.TypeOf(GetType(Integer)))
        End Sub

        <Test()>
        Public Sub InstanceOfTests()
            ' Classic syntax
            Assert.IsInstanceOf(GetType(String), "Hello")
            Assert.IsNotInstanceOf(GetType(String), 5)

            ' Helper syntax
            Assert.That("Hello", [Is].InstanceOf(GetType(String)))
            Assert.That(5, [Is].Not.InstanceOf(GetType(String)))
        End Sub

        <Test()>
        Public Sub AssignableFromTypeTests()
            ' Classic syntax
            Assert.IsAssignableFrom(GetType(String), "Hello")
            Assert.IsNotAssignableFrom(GetType(String), 5)

            ' Helper syntax
            Assert.That("Hello", [Is].AssignableFrom(GetType(String)))
            Assert.That(5, [Is].Not.AssignableFrom(GetType(String)))
        End Sub
#End Region

#Region "StringConstraintTests"
        <Test()>
        Public Sub SubstringTests()
            Dim phrase As String = "Hello World!"
            Dim array As String() = New String() {"abc", "bad", "dba"}

            ' Classic Syntax
            StringAssert.Contains("World", phrase)

            ' Helper syntax
            Assert.That(phrase, Does.Contain("World"))
            ' Only available using new syntax
            Assert.That(phrase, Does.Not.Contains("goodbye"))
            Assert.That(phrase, Does.Contain("WORLD").IgnoreCase)
            Assert.That(phrase, Does.Not.Contains("BYE").IgnoreCase)
            Assert.That(array, [Is].All.Contains("b"))
        End Sub

        <Test()>
        Public Sub StartsWithTests()
            Dim phrase As String = "Hello World!"
            Dim greetings As String() = New String() {"Hello!", "Hi!", "Hola!"}

            ' Classic syntax
            StringAssert.StartsWith("Hello", phrase)

            ' Helper syntax
            Assert.That(phrase, Does.StartWith("Hello"))
            ' Only available using new syntax
            Assert.That(phrase, Does.Not.StartsWith("Hi!"))
            Assert.That(phrase, Does.StartWith("HeLLo").IgnoreCase)
            Assert.That(phrase, Does.Not.StartsWith("HI").IgnoreCase)
            Assert.That(greetings, [Is].All.StartsWith("h").IgnoreCase)
        End Sub

        <Test()>
        Public Sub EndsWithTests()
            Dim phrase As String = "Hello World!"
            Dim greetings As String() = New String() {"Hello!", "Hi!", "Hola!"}

            ' Classic Syntax
            StringAssert.EndsWith("!", phrase)

            ' Helper syntax
            Assert.That(phrase, Does.EndWith("!"))
            ' Only available using new syntax
            Assert.That(phrase, Does.Not.EndWith("?"))
            Assert.That(phrase, Does.EndWith("WORLD!").IgnoreCase)
            Assert.That(greetings, [Is].All.EndsWith("!"))
        End Sub

        <Test()>
        Public Sub EqualIgnoringCaseTests()

            Dim phrase As String = "Hello World!"
            Dim array1 As String() = New String() {"Hello", "World"}
            Dim array2 As String() = New String() {"HELLO", "WORLD"}
            Dim array3 As String() = New String() {"HELLO", "Hello", "hello"}

            ' Classic syntax
            StringAssert.AreEqualIgnoringCase("hello world!", phrase)

            ' Helper syntax
            Assert.That(phrase, [Is].EqualTo("hello world!").IgnoreCase)
            'Only available using new syntax
            Assert.That(phrase, [Is].Not.EqualTo("goodbye world!").IgnoreCase)
            Assert.That(array1, [Is].EqualTo(array2).IgnoreCase)
            Assert.That(array3, [Is].All.EqualTo("hello").IgnoreCase)
        End Sub

        <Test()>
        Public Sub RegularExpressionTests()
            Dim phrase As String = "Now is the time for all good men to come to the aid of their country."
            Dim quotes As String() = New String() {"Never say never", "It's never too late", "Nevermore!"}

            ' Classic syntax
            StringAssert.IsMatch("all good men", phrase)
            StringAssert.IsMatch("Now.*come", phrase)

            ' Helper syntax
            Assert.That(phrase, Does.Match("all good men"))
            Assert.That(phrase, Does.Match("Now.*come"))
            ' Only available using new syntax
            Assert.That(phrase, Does.Not.Match("all.*men.*good"))
            Assert.That(quotes, [Is].All.Matches("never").IgnoreCase)
        End Sub
#End Region

#Region "Equality Tests"
        <Test()>
        Public Sub EqualityTests()

            Dim i3 As Integer() = {1, 2, 3}
            Dim d3 As Double() = {1.0, 2.0, 3.0}
            Dim iunequal As Integer() = {1, 3, 2}

            ' Classic Syntax
            Assert.AreEqual(4, 2 + 2)
            Assert.AreEqual(i3, d3)
            Assert.AreNotEqual(5, 2 + 2)
            Assert.AreNotEqual(i3, iunequal)

            ' Helper syntax
            Assert.That(2 + 2, [Is].EqualTo(4))
            Assert.That(2 + 2 = 4)
            Assert.That(i3, [Is].EqualTo(d3))
            Assert.That(2 + 2, [Is].Not.EqualTo(5))
            Assert.That(i3, [Is].Not.EqualTo(iunequal))
        End Sub

        <Test()>
        Public Sub EqualityTestsWithTolerance()
            ' CLassic syntax
            Assert.AreEqual(5.0R, 4.99R, 0.05R)
            Assert.AreEqual(5.0F, 4.99F, 0.05F)

            ' Helper syntax
            Assert.That(4.99R, [Is].EqualTo(5.0R).Within(0.05R))
            Assert.That(4D, [Is].Not.EqualTo(5D).Within(0.5D))
            Assert.That(4.99F, [Is].EqualTo(5.0F).Within(0.05F))
            Assert.That(4.99D, [Is].EqualTo(5D).Within(0.05D))
            Assert.That(499, [Is].EqualTo(500).Within(5))
            Assert.That(4999999999L, [Is].EqualTo(5000000000L).Within(5L))
        End Sub

        <Test()>
        Public Sub EqualityTestsWithTolerance_MixedFloatAndDouble()
            ' Bug Fix 1743844
            Assert.That(2.20492R, [Is].EqualTo(2.2R).Within(0.01F),
                "Double actual, Double expected, Single tolerance")
            Assert.That(2.20492R, [Is].EqualTo(2.2F).Within(0.01R),
                "Double actual, Single expected, Double tolerance")
            Assert.That(2.20492R, [Is].EqualTo(2.2F).Within(0.01F),
                "Double actual, Single expected, Single tolerance")
            Assert.That(2.20492F, [Is].EqualTo(2.2F).Within(0.01R),
                "Single actual, Single expected, Double tolerance")
            Assert.That(2.20492F, [Is].EqualTo(2.2R).Within(0.01R),
                "Single actual, Double expected, Double tolerance")
            Assert.That(2.20492F, [Is].EqualTo(2.2R).Within(0.01F),
                "Single actual, Double expected, Single tolerance")
        End Sub

        <Test()>
        Public Sub EqualityTestsWithTolerance_MixingTypesGenerally()
            ' Extending tolerance to all numeric types
            Assert.That(202.0R, [Is].EqualTo(200.0R).Within(2),
                "Double actual, Double expected, int tolerance")
            Assert.That(4.87D, [Is].EqualTo(5).Within(0.25R),
                "Decimal actual, int expected, Double tolerance")
            Assert.That(4.87D, [Is].EqualTo(5L).Within(1),
                "Decimal actual, long expected, int tolerance")
            Assert.That(487, [Is].EqualTo(500).Within(25),
                "int actual, int expected, int tolerance")
            Assert.That(487L, [Is].EqualTo(500).Within(25),
                "long actual, int expected, int tolerance")
        End Sub
#End Region

#Region "Comparison Tests"
        <Test()>
        Public Sub ComparisonTests()
            ' Classic Syntax
            Assert.Greater(7, 3)
            Assert.GreaterOrEqual(7, 3)
            Assert.GreaterOrEqual(7, 7)

            ' Helper syntax
            Assert.That(7, [Is].GreaterThan(3))
            Assert.That(7, [Is].GreaterThanOrEqualTo(3))
            Assert.That(7, [Is].AtLeast(3))
            Assert.That(7, [Is].GreaterThanOrEqualTo(7))
            Assert.That(7, [Is].AtLeast(7))

            ' Classic syntax
            Assert.Less(3, 7)
            Assert.LessOrEqual(3, 7)
            Assert.LessOrEqual(3, 3)

            ' Helper syntax
            Assert.That(3, [Is].LessThan(7))
            Assert.That(3, [Is].LessThanOrEqualTo(7))
            Assert.That(3, [Is].AtMost(7))
            Assert.That(3, [Is].LessThanOrEqualTo(3))
            Assert.That(3, [Is].AtMost(3))
        End Sub
#End Region

#Region "Collection Tests"
        <Test()>
        Public Sub AllItemsTests()

            Dim ints As Object() = {1, 2, 3, 4}
            Dim doubles As Object() = {0.99, 2.1, 3.0, 4.05}
            Dim strings As Object() = {"abc", "bad", "cab", "bad", "dad"}

            ' Classic syntax
            CollectionAssert.AllItemsAreNotNull(ints)
            CollectionAssert.AllItemsAreInstancesOfType(ints, GetType(Integer))
            CollectionAssert.AllItemsAreInstancesOfType(strings, GetType(String))
            CollectionAssert.AllItemsAreUnique(ints)

            ' Helper syntax
            Assert.That(ints, [Is].All.Not.Null)
            Assert.That(ints, Has.None.Null)
            Assert.That(ints, [Is].All.InstanceOf(GetType(Integer)))
            Assert.That(ints, Has.All.InstanceOf(GetType(Integer)))
            Assert.That(strings, [Is].All.InstanceOf(GetType(String)))
            Assert.That(strings, Has.All.InstanceOf(GetType(String)))
            Assert.That(ints, Iz.Unique)
            ' Only available using new syntax
            Assert.That(strings, [Is].Not.Unique)
            Assert.That(ints, [Is].All.GreaterThan(0))
            Assert.That(ints, Has.All.GreaterThan(0))
            Assert.That(ints, Has.None.LessThanOrEqualTo(0))
            Assert.That(strings, [Is].All.Contains("a"))
            Assert.That(strings, Has.All.Contains("a"))
            Assert.That(strings, Has.Some.StartsWith("ba"))
            Assert.That(strings, Has.Some.Property("Length").EqualTo(3))
            Assert.That(strings, Has.Some.StartsWith("BA").IgnoreCase)
            Assert.That(doubles, Has.Some.EqualTo(1.0).Within(0.05))
        End Sub

        <Test()>
        Public Sub SomeItemsTests()

            Dim mixed As Object() = {1, 2, "3", Nothing, "four", 100}
            Dim strings As Object() = {"abc", "bad", "cab", "bad", "dad"}

            ' Not available using the classic syntax

            ' Helper syntax
            Assert.That(mixed, Has.Some.Null)
            Assert.That(mixed, Has.Some.InstanceOf(GetType(Integer)))
            Assert.That(mixed, Has.Some.InstanceOf(GetType(String)))
            Assert.That(strings, Has.Some.StartsWith("ba"))
            Assert.That(strings, Has.Some.Not.StartsWith("ba"))
        End Sub

        <Test()>
        Public Sub NoItemsTests()

            Dim ints As Object() = {1, 2, 3, 4, 5}
            Dim strings As Object() = {"abc", "bad", "cab", "bad", "dad"}

            ' Not available using the classic syntax

            ' Helper syntax
            Assert.That(ints, Has.None.Null)
            Assert.That(ints, Has.None.InstanceOf(GetType(String)))
            Assert.That(ints, Has.None.GreaterThan(99))
            Assert.That(strings, Has.None.StartsWith("qu"))
        End Sub

        <Test()>
        Public Sub CollectionContainsTests()

            Dim iarray As Integer() = {1, 2, 3}
            Dim sarray As String() = {"a", "b", "c"}

            ' Classic syntax
            Assert.Contains(3, iarray)
            Assert.Contains("b", sarray)
            CollectionAssert.Contains(iarray, 3)
            CollectionAssert.Contains(sarray, "b")
            CollectionAssert.DoesNotContain(sarray, "x")
            ' Showing that Contains uses NUnit equality
            CollectionAssert.Contains(iarray, 1.0R)

            ' Helper syntax
            Assert.That(iarray, Has.Member(3))
            Assert.That(sarray, Has.Member("b"))
            Assert.That(sarray, Has.No.Member("x"))
            ' Showing that Contains uses NUnit equality
            Assert.That(iarray, Has.Member(1.0R))

            ' Only available using the new syntax
            ' Note that EqualTo and SameAs do NOT give
            ' identical results to Contains because 
            ' Contains uses Object.Equals()
            Assert.That(iarray, Has.Some.EqualTo(3))
            Assert.That(iarray, Has.Member(3))
            Assert.That(sarray, Has.Some.EqualTo("b"))
            Assert.That(sarray, Has.None.EqualTo("x"))
            Assert.That(iarray, Has.None.SameAs(1.0R))
            Assert.That(iarray, Has.All.LessThan(10))
            Assert.That(sarray, Has.All.Length.EqualTo(1))
            Assert.That(sarray, Has.None.Property("Length").GreaterThan(3))
        End Sub

        <Test()>
        Public Sub CollectionEquivalenceTests()

            Dim ints1to5 As Integer() = {1, 2, 3, 4, 5}
            Dim twothrees As Integer() = {1, 2, 3, 3, 4, 5}
            Dim twofours As Integer() = {1, 2, 3, 4, 4, 5}

            ' Classic syntax
            CollectionAssert.AreEquivalent(New Integer() {2, 1, 4, 3, 5}, ints1to5)
            CollectionAssert.AreNotEquivalent(New Integer() {2, 2, 4, 3, 5}, ints1to5)
            CollectionAssert.AreNotEquivalent(New Integer() {2, 4, 3, 5}, ints1to5)
            CollectionAssert.AreNotEquivalent(New Integer() {2, 2, 1, 1, 4, 3, 5}, ints1to5)
            CollectionAssert.AreNotEquivalent(twothrees, twofours)

            ' Helper syntax
            Assert.That(New Integer() {2, 1, 4, 3, 5}, Iz.EquivalentTo(ints1to5))
            Assert.That(New Integer() {2, 2, 4, 3, 5}, [Is].Not.EquivalentTo(ints1to5))
            Assert.That(New Integer() {2, 4, 3, 5}, [Is].Not.EquivalentTo(ints1to5))
            Assert.That(New Integer() {2, 2, 1, 1, 4, 3, 5}, [Is].Not.EquivalentTo(ints1to5))
            Assert.That(twothrees, [Is].Not.EquivalentTo(twofours))
        End Sub

        <Test()>
        Public Sub SubsetTests()

            Dim ints1to5 As Integer() = {1, 2, 3, 4, 5}

            ' Classic syntax
            CollectionAssert.IsSubsetOf(New Integer() {1, 3, 5}, ints1to5)
            CollectionAssert.IsSubsetOf(New Integer() {1, 2, 3, 4, 5}, ints1to5)
            CollectionAssert.IsNotSubsetOf(New Integer() {2, 4, 6}, ints1to5)
            CollectionAssert.IsNotSubsetOf(New Integer() {1, 2, 2, 2, 5}, ints1to5)

            ' Helper syntax
            Assert.That(New Integer() {1, 3, 5}, [Is].SubsetOf(ints1to5))
            Assert.That(New Integer() {1, 2, 3, 4, 5}, [Is].SubsetOf(ints1to5))
            Assert.That(New Integer() {2, 4, 6}, [Is].Not.SubsetOf(ints1to5))
        End Sub
#End Region

#Region "Property Tests"
        <Test()>
        Public Sub PropertyTests()

            Dim array As String() = {"abc", "bca", "xyz", "qrs"}
            Dim array2 As String() = {"a", "ab", "abc"}
            Dim list As New ArrayList(array)

            ' Not available using the classic syntax

            ' Helper syntax
            Assert.That(list, Has.Property("Count"))
            Assert.That(list, Has.No.Property("Length"))

            Assert.That("Hello", Has.Length.EqualTo(5))
            Assert.That("Hello", Has.Property("Length").EqualTo(5))
            Assert.That("Hello", Has.Property("Length").GreaterThan(3))

            Assert.That(array, Has.Property("Length").EqualTo(4))
            Assert.That(array, Has.Length.EqualTo(4))
            Assert.That(array, Has.Property("Length").LessThan(10))

            Assert.That(array, Has.All.Property("Length").EqualTo(3))
            Assert.That(array, Has.All.Length.EqualTo(3))
            Assert.That(array, [Is].All.Length.EqualTo(3))
            Assert.That(array, Has.All.Property("Length").EqualTo(3))
            Assert.That(array, [Is].All.Property("Length").EqualTo(3))

            Assert.That(array2, [Is].Not.Property("Length").EqualTo(4))
            Assert.That(array2, [Is].Not.Length.EqualTo(4))
            Assert.That(array2, Has.No.Property("Length").GreaterThan(3))
        End Sub
#End Region

#Region "Not Tests"
        <Test()>
        Public Sub NotTests()
            ' Not available using the classic syntax

            ' Helper syntax
            Assert.That(42, [Is].Not.Null)
            Assert.That(42, [Is].Not.True)
            Assert.That(42, [Is].Not.False)
            Assert.That(2.5, [Is].Not.NaN)
            Assert.That(2 + 2, [Is].Not.EqualTo(3))
            Assert.That(2 + 2, [Is].Not.Not.EqualTo(4))
            Assert.That(2 + 2, [Is].Not.Not.Not.EqualTo(5))
        End Sub
#End Region

    End Class

End Namespace
