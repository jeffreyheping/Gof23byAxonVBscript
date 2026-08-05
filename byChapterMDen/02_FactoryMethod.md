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
- This pattern maps naturally to AxonASP. The interface mechanism solves the core polymorphism problem with no significant trade-offs.
---
