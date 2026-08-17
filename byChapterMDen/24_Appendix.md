## Appendix: 23 Pattern Implementation Status

This book covers four runtimes. Classic VBScript code runs on three engines (cscript, AxonASP classic mode, ASPPY); Axon VBScript code runs on AxonASP; VB.NET code runs on .NET. The table below compares implementation status across the two VBScript syntax systems, with VB.NET as a syntactically complete baseline (all 23 patterns implemented idiomatically, no compromises needed).

### I. Already Idiomatic in Classic (3)

These three patterns are already idiomatically implemented in classic VBScript without needing interface/event extensions. AxonASP adds strongly-typed (`As Type`/UDT) versions for improved type safety and IDE support, but doesn't change the pattern structure — so they remain classified as "already idiomatic in classic".

| # | Pattern | Why idiomatic | AxonASP strong-typing version |
|:---:|---|---|---|
| 7 | Facade | Composes calls to subsystems, no inheritance or interfaces needed | Subsystem references `As CPU`/`As Memory`/`As HardDrive`, parameters `As Long`/`As String` |
| 12 | Flyweight | `Dictionary` for object caching is already optimal | `TreeType` fields and `GetTreeType` return value strongly typed; `m_Types` stays Variant due to COM limitations |
| 22 | Memento | `Class` + `Property` for state snapshots is already idiomatic | UDT `EditorMemento` + `As String`/`As Integer`, value-copy prevents reference leaks |

### II. AxonASP Fully Solves the Core Problem (18)

| # | Pattern | Feature used | What it solves |
|:---:|---|---|---|
| 2 | Factory Method | `Implements` | `IAnimal` interface constrains product contract |
| 3 | Abstract Factory | `Implements` | `IGUIFactory` + product family interface constraints |
| 4 | Builder | `Implements` | `IBuilder` constrains contract; VBScript lacks chaining syntax sugar, requires line-by-line calls |
| 5 | Prototype | `Implements` | `ICloneable` guarantees contract; deep copy must be manual (same in Go) |
| 6 | Proxy | `Implements` | `IImage` unifies proxy and real object |
| 8 | Adapter | `Implements` | `IPrinter` enforces adapter contract |
| 9 | Bridge | `Implements` | `IRenderer` separates abstraction from implementation |
| 10 | Composite | `Implements` | `IComponent` unifies leaves and branches |
| 11 | Decorator | `Implements` | `ICoffee` enables transparent decorator/component substitution |
| 13 | Strategy | `Implements` | `ISortStrategy` constrains algorithm contract |
| 14 | Observer | `Event` | `RaiseEvent` auto-notifies, no more manual array management |
| 15 | Template Method | `Implements` | `IExtractor` injects variable steps |
| 16 | Iterator | `For Each` custom collection iteration | Pattern disappears entirely — `[DispId(-4)]` forwards built-in Collection enumerator |
| 17 | Chain of Responsibility | `Implements` | `IHandler` chain node contract, `LogLevel` enum replaces magic numbers |
| 18 | Command | `Implements` | `ICommand` command contract |
| 19 | State | `Implements` | `IState` constrains both `Handle` behavior and `NextState` transition contract; state transitions delegated to each state class |
| 20 | Mediator | `Implements` | `IMediator`/`IColleague` constrain contracts; centralized interaction is a pattern characteristic, not a language limitation |
| 23 | Interpreter | `Implements` | `IExpression` recursive type safety |

### III. AxonASP Improved but Still Has Residual Defects (2)

| # | Pattern | What Axon solves | Missing syntax | Go's alternative |
|:---:|---|---|---|---|
| 1 | Singleton | `Static` eliminates global variables | Private constructor | Package-private (lowercase unexported) |
| 21 | Visitor | `IVisitor`/`IElement` constrain contracts | Method overloading or double dispatch | Go uses type switch for double dispatch |

### IV. VB.NET Baseline

VB.NET, as a syntactically complete OOP language, idiomatically implements all 23 patterns with no compromises. The "residual defects" in the table above all have solutions in VB.NET:

| Residual defect | VB.NET solution |
|---------|------------|
| Private constructor (Singleton) | `Private Sub New()` enforces compile-time prevention of external instantiation |
| Method overloading / double dispatch (Visitor) | `Overloads` method overloading + runtime type dispatch |

### V. Missing Syntax Priorities for GoF-23

Filtered by "does it help the idiomatic implementation of the 23 design patterns", the syntax still missing from AxonASP:

| Priority | Missing syntax | Affected patterns | Go's approach | Rationale |
|:---:|---|---|---|---|
| P0 | Code reuse mechanism (inheritance or embedding) | Proxy, Bridge, Composite, Chain of Responsibility, Template Method | struct embedding | Doesn't affect functionality — all patterns can already be implemented with interfaces + composition. But manual delegate forwarding creates lots of boilerplate; Go's embedding solves this directly |
| P1 | Method overloading / double dispatch | Visitor only | type switch | Narrowest scope, but Visitor pattern genuinely needs double dispatch to be idiomatic |
| P2 | Access control modifiers | Singleton, Memento | First-letter casing controls visibility | Small impact; Singleton's private constructor is the only hard requirement |

> **Resolved**: `For Each` custom collection iteration (originally P0) has been implemented by AxonASP — the Iterator pattern disappears entirely.

### Summary

| Status | Classic VBScript | AxonASP | VB.NET |
|------|:---:|:---:|:---:|
| Already idiomatic in classic | 3 | 3 | 3 |
| Fully solved | 0 | 18 | 20 |
| Improved but residual | 0 | 2 | 0 |
| **Idiomatic total** | **3** | **21** | **23** |

AxonASP raises the idiomatic implementation count from **3/23** to **21/23**. The 18 "fully solved" patterns achieve compile-time type-safe polymorphic dispatch and native traversal through `Implements` interfaces, `Static` static variables, `Event` mechanism, and `For Each` custom collection iteration — no helper classes or fully-qualified names needed. VB.NET, as the baseline, idiomatically implements all 23 patterns — confirming the decisive role of "syntactic completeness" in pattern implementation.

The residual defects in the remaining 2 patterns fall into two categories when compared against Go:

- **Go also lacks it but has alternatives**: Method overloading / double dispatch (Go uses type switch). AxonASP currently substitutes with interfaces + manual branching — functionally equivalent but more boilerplate.
- **Go has it but AxonASP doesn't**: Access control (Go's first-letter casing). Singleton's private constructor is the only hard requirement.

Go lacks both "inheritance" and "method overloading" — two features widely considered "essential for OOP" — yet still idiomatically implements all 23 GoF patterns. This shows these features aren't necessary for design patterns — interfaces + composition are the core. AxonASP already has interfaces and composition; the direction is right.

### VI. AxonASP Features Used in This Book

AxonASP introduces many modern features to VBScript; this book uses only 8 of them. Filtered by "does it help the idiomatic implementation of the 23 design patterns", 4 are core and 4 are nice-to-have.

**Core (passed the GoF-23 filter):**

| Feature | Where used | Why core |
|---|---|---|
| `Implements` interfaces | Polymorphic dispatch in 17 patterns | Without it, Factory, Proxy, Adapter, Strategy etc. can only rely on method name conventions — not idiomatic |
| `Event`/`RaiseEvent`/`WithEvents` | Chapter 14 Observer | Without it, Observer must manually maintain an array — not idiomatic |
| `Static` static variables | Chapter 1 Singleton | Without it, Singleton must use global variables, breaking encapsulation |
| `For Each` custom collection iteration + `[DispId(-4)]` | Chapter 16 Iterator | Without it, Iterator needs hand-written `HasNext`/`NextItem` classes; with it, the Iterator pattern disappears entirely |

**Nice-to-have (didn't pass the GoF-23 filter):**

| Feature | Where used | Why it doesn't pass the filter | Value |
|---|---|---|---|
| Strong typing (`As Type`) | 17 "fully solved" patterns (except Iterator #16 which depends on built-in Collection COM object and can't be fully annotated) + 3 "already idiomatic" patterns' subsystem references, fields, and parameters — 20 patterns total | All 23 patterns can be implemented with Variant; type safety isn't a pattern idiomaticity issue | Eliminates entire classes of runtime type errors; foundation for IDE intellisense |
| UDT (`Type`) | Chapter 22 Memento (`EditorMemento`) | Memento is already idiomatic with `Class` + `Property`; UDT is just more elegant | Value semantics naturally prevent reference leaks, reduce boilerplate |
| `Enum` enumeration | Chapter 17 Chain of Responsibility (`LogLevel`) | Can be replaced with `Const`, functionally equivalent | Improved readability, IDE autocomplete support |
| `IsNot` operator | Chapter 17 Chain of Responsibility (`m_Next IsNot Nothing`) | `Not x Is Nothing` is functionally equivalent; `IsNot` is just clearer | Eliminates `Not ... Is Nothing` precedence ambiguity, guard clauses more readable |

Without these 4 core features, 18 of the 23 patterns can't be idiomatically implemented; without the 4 nice-to-have features, all 23 patterns are still idiomatic — but the code is safer, cleaner, and more maintainable. Strong typing already covers 17 "fully solved" patterns (Iterator #16 can't annotate `m_Items` etc. due to COM interop limitations) plus 3 "already idiomatic" patterns, totaling 20; the remaining Chapter 16 Iterator retains Variant due to COM interop constraints. This shows "nice-to-have" doesn't mean "insignificant" — it makes already-idiomatic code safer and more maintainable. If AxonASP introduces more general language improvements in the future (like `AndAlso`/`OrElse` short-circuit evaluation, parameterized constructors, etc.), the GoF-23 filter can similarly be used to evaluate their priority — independent of the classification here.

---

## Appendix: 23 Patterns at a Glance

| # | Pattern | Category | One-line summary |
|:---:|---|---|---|
| 1 | **Singleton** | Creational | Only one instance globally |
| 2 | **Factory Method** | Creational | Delegates creation to subclasses/factories |
| 3 | **Abstract Factory** | Creational | Swap factory, swap entire product family |
| 4 | **Builder** | Creational | Step-by-step construction of complex objects |
| 5 | **Prototype** | Creational | Creates new objects by copying |
| 6 | **Proxy** | Structural | Provides a surrogate/placeholder for an object |
| 7 | **Facade** | Structural | Simple interface for a complex system |
| 8 | **Adapter** | Structural | Converts incompatible interface to target interface |
| 9 | **Bridge** | Structural | Separates abstraction from implementation |
| 10 | **Composite** | Structural | Treats individual and composite objects uniformly |
| 11 | **Decorator** | Structural | Dynamically adds functionality without changing original class |
| 12 | **Flyweight** | Structural | Shares fine-grained objects to save memory |
| 13 | **Strategy** | Behavioral | Encapsulates algorithms, swappable at runtime |
| 14 | **Observer** | Behavioral | Notifies all subscribers of state changes |
| 15 | **Template Method** | Behavioral | Defines algorithm skeleton, subclasses fill in steps |
| 16 | **Iterator** | Behavioral | Sequential access to collection elements without exposing internals |
| 17 | **Chain of Responsibility** | Behavioral | Passes request along the chain until handled |
| 18 | **Command** | Behavioral | Encapsulates a request as an object |
| 19 | **State** | Behavioral | Behavior changes with internal state |
| 20 | **Mediator** | Behavioral | Encapsulates object interactions via a mediator |
| 21 | **Visitor** | Behavioral | Encapsulates operations in a visitor |
| 22 | **Memento** | Behavioral | Saves and restores object state |
| 23 | **Interpreter** | Behavioral | Builds an interpreter for a language |
