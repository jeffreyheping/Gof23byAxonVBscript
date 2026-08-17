## Chapter 2: Factory Method

**Core idea**: Encapsulate the "which object to create" decision inside a factory.

**Example**: Dog and Cat are two product classes. AnimalFactory decides which animal to create based on a type string. The caller doesn't need to know the concrete class names.

### Classic VBScript Version

```vbscript
' Product: Dog
Class Dog
    ' Make the dog speak
    Public Function Speak
        Response.Write "Woof"
    End Function
End Class

' Product: Cat
Class Cat
    ' Make the cat speak
    Public Function Speak
        Response.Write "Meow"
    End Function
End Class

' Factory: creates the right animal based on a type string
' animalType: "dog" or "cat"
' Returns: a Dog or Cat instance
Class AnimalFactory
    Public Function CreateAnimal(animalType)
        Select Case LCase(animalType)
            Case "dog"
                Set CreateAnimal = New Dog
            Case "cat"
                Set CreateAnimal = New Cat
            Case Else
                Set CreateAnimal = Nothing
        End Select
    End Function
End Class

' Demo: create objects via the factory, not directly with New
Dim factory
Set factory = New AnimalFactory
Set myPet = factory.CreateAnimal("dog")
myPet.Speak   ' Woof
```

**Classic VBScript trade-offs**:
- **No inheritance, no abstract methods**: The classic Factory Method relies on an abstract Creator with subclasses overriding the factory method. VBScript has no inheritance, so we put `Select Case` inside the factory class — every new product requires changing the factory code, violating the Open-Closed Principle.
- **No interface constraint**: Dog and Cat have no `IAnimal` interface guaranteeing they both have `Speak`. It's up to the developer to stay consistent. Wrong types are only caught at runtime.

### Axon VBScript Version (supports Implements)

```vbscript
' Product interface
Class IAnimal
    Public Function Speak
    End Function
End Class

' Factory interface
Class IFactory
    Public Function CreateAnimal As IAnimal
    End Function
End Class

' Concrete product: Dog
Class Dog
    Implements IAnimal
    Public Function IAnimal_Speak
        Response.Write "Woof"
    End Function
End Class

' Concrete product: Cat
Class Cat
    Implements IAnimal
    Public Function IAnimal_Speak
        Response.Write "Meow"
    End Function
End Class

' Concrete factory: creates Dogs
Class DogFactory
    Implements IFactory
    Public Function IFactory_CreateAnimal As IAnimal
        Set IFactory_CreateAnimal = New Dog
    End Function
End Class

' Concrete factory: creates Cats
Class CatFactory
    Implements IFactory
    Public Function IFactory_CreateAnimal As IAnimal
        Set IFactory_CreateAnimal = New Cat
    End Function
End Class

' Demo: create animals through the factory interface; caller doesn't depend on concrete classes
Dim factory As IFactory
Dim myPet As IAnimal
Set factory = New DogFactory
Set myPet = factory.CreateAnimal
myPet.Speak   ' Woof
```

**Axon VBScript trade-offs**:
- This pattern maps naturally to AxonASP. The interface mechanism solves the core polymorphism problem. Remaining gap: **missing code reuse mechanism (inheritance)**. Classic Factory Method requires "abstract Creator base class + subclass overriding FactoryMethod". AxonASP currently can only use two separate interfaces (`IFactory`/`IAnimal`) as a substitute, lacking an abstract base class to constrain and reuse Creator's common logic.

### VB.NET Version (syntactically complete baseline)

VB.NET has `MustInherit` (abstract class) + `MustOverride` (abstract method) + `Inherits` (inheritance), enabling textbook Factory Method — abstract Creator defines the skeleton, concrete Factory subclasses override the factory method.

```vbnet
' ① Abstract product base class: MustInherit prevents direct instantiation, MustOverride forces subclasses to implement Speak
Public MustInherit Class Animal
    Public MustOverride Sub Speak()
End Class

' ② Abstract factory base class: defines CreateAnimal skeleton, no concrete creation logic — left to subclasses
Public MustInherit Class AnimalFactory
    Public MustOverride Function CreateAnimal() As Animal
End Class

' ③ Concrete products
Public Class Dog
    Inherits Animal
    Public Overrides Sub Speak()
        Console.WriteLine("Woof")
    End Sub
End Class

Public Class Cat
    Inherits Animal
    Public Overrides Sub Speak()
        Console.WriteLine("Meow")
    End Sub
End Class

' ④ Concrete factories: each Factory subclass handles one product, naturally following Open-Closed Principle
Public Class DogFactory
    Inherits AnimalFactory
    Public Overrides Function CreateAnimal() As Animal
        Return New Dog()
    End Function
End Class

Public Class CatFactory
    Inherits AnimalFactory
    Public Overrides Function CreateAnimal() As Animal
        Return New Cat()
    End Function
End Class

' Demo: caller depends only on abstract base classes, completely independent of concrete Factory/Product
Dim factory As AnimalFactory = New DogFactory()
Dim myPet As Animal = factory.CreateAnimal()
myPet.Speak()   ' Woof
```

**VB.NET version notes**:
- **Real abstract base classes**: `MustInherit Class Animal` prevents `New Animal()`, `MustOverride Speak` forces all subclasses to implement — compile-time check, missing implementation causes immediate error. Compare with Axon version: `IAnimal` interface can constrain `Speak` exists, but cannot prevent external `New IAnimal` (VBScript classes cannot be marked as abstract/non-instantiable).
- **Code reuse through inheritance**: If you later want to add `Name` property and `Born()` method to all products, just add once in `Animal` base class, all subclasses inherit automatically. Axon version must manually write Dog and Cat each with their own copy — more products means more duplication.
- **Open-Closed Principle truly realized**: Adding new products only needs a new `*Product` class + a `*Factory` class, **zero changes to existing code**. Axon version also satisfies Open-Closed, but without an abstract base class as "contract anchor", all factories and products are isolated parallel classes, lacking architectural-level constraint.
- **No `Set`/`Let` distinction**: VB.NET object assignment uses `=` directly, no need to remember `Set` for objects, `Let` for value types.

**Three-version comparison**:

| Dimension | Classic VBScript | Axon VBScript | VB.NET |
|------|--------------|---------------|--------|
| Product constraint | Method name convention (easy to miss) | `IAnimal` interface constrains `Speak` | `MustInherit` + `MustOverride` compile-time enforced |
| Factory implementation | Single factory internal `Select Case` (change one, change all) | One Factory class per product + `IFactory` interface | `MustInherit` abstract base class + subclass `Overrides` |
| Open-Closed | No (new product changes Factory's Case chain) | Yes (new product = two new classes) | Yes (clearer inheritance, base class as contract anchor) |
| Code reuse | None | None (parallel product/factory classes, no inheritance) | Base class fields/methods automatically inherited by all subclasses |
| Object assignment | `Set a = New X` | `Set a = New X` | Direct `a = New X()` |
---
