## Appendix: 23 Patterns Implementation Status Overview

AxonASP introduces modern extensions to VBScript: `Implements` interface polymorphism, `Static` static variables, `Event`/`RaiseEvent`/`WithEvents` event mechanism, and more. Below is an implementation status comparison for all 23 patterns across classic VBScript and AxonASP.

### I. Classic Already Idiomatic (3 patterns)

These three patterns need no interface/event extensions to be idiomatically implemented in classic VBScript. AxonASP adds strongly-typed (`As Type`/UDT) versions for improved type safety and IDE support, but doesn't change the pattern structure — so they remain classified as "classic already idiomatic".

| # | Pattern | Idiomatic reason | AxonASP strongly-typed version |
|:---:|---|---|---|
| 7 | Facade | Composes subsystem calls; no inheritance or interfaces needed | Subsystem refs `As CPU`/`As Memory`/`As HardDrive`, params `As Long`/`As String` |
| 12 | Flyweight | `Dictionary` for object caching is already optimal | `TreeType` fields and `GetTreeType` return type strongly typed; `m_Types` stays Variant due to COM limits |
| 22 | Memento | `Class` + `Property` for state snapshots is already idiomatic | UDT `EditorMemento` + `As String`/`As Integer`; value copy prevents reference leaks |

### II. AxonASP Fully Resolves Core Pain Point (18 patterns)

| # | Pattern | Feature used | What it solves |
|:---:|---|---|---|
| 2 | Factory Method | `Implements` | `IAnimal` interface constrains product contract |
| 3 | Abstract Factory | `Implements` | `IGUIFactory` + product family interfaces |
| 4 | Builder | `Implements` | `IBuilder` constrains contract; VBScript lacks chained-call sugar, requires line-by-line calls |
| 5 | Prototype | `Implements` | `ICloneable` guarantees contract; deep copy still manual (same in Go) |
| 6 | Proxy | `Implements` | `IImage` unifies proxy and real object |
| 8 | Adapter | `Implements` | `IPrinter` enforces adapter contract |
| 9 | Bridge | `Implements` | `IRenderer` separates abstraction from implementation |
| 10 | Composite | `Implements` | `IComponent` unifies leaves and branches |
| 11 | Decorator | `Implements` | `ICoffee` enables transparent decorator/component substitution |
| 13 | Strategy | `Implements` | `ISortStrategy` constrains algorithm contract |
| 14 | Observer | `Event` | `RaiseEvent` auto-notifies; no more manual arrays |
| 15 | Template Method | `Implements` | `IExtractor` injects the variable step |
| 16 | Iterator | `For Each` custom collection | Pattern disappears entirely — `[DispId(-4)]` forwards built-in Collection enumerator |
| 17 | Chain of Responsibility | `Implements` | `IHandler` chain node contract; `LogLevel` enum replaces magic strings |
| 18 | Command | `Implements` | `ICommand` command contract |
| 19 | State | `Implements` | `IState` constrains both `Handle` behavior and `NextState` transition; transition logic pushed to state classes |
| 20 | Mediator | `Implements` | `IMediator`/`IColleague` constrain contracts; centralized interaction is the pattern's nature, not a language flaw |
| 23 | Interpreter | `Implements` | `IExpression` recursive type safety |

### III. AxonASP Improves but Residual Gaps Remain (2 patterns)

| # | Pattern | What Axon solves | Missing syntax | Go's alternative |
|:---:|---|---|---|---|
| 1 | Singleton | `Static` eliminates global variable | Private constructor | Package-private (lowercase unexported) |
| 21 | Visitor | `IVisitor`/`IElement` constrain contracts | Method overloading / double dispatch | Go uses type switch for double dispatch |

### IV. Missing Syntax Priorities for GoF-23

Filtered by "does it help make the 23 patterns more idiomatic", AxonASP's remaining syntax gaps are:

| Priority | Missing syntax | Affected patterns | Go's approach | Rationale |
|:---:|---|---|---|---|
| P0 | Code reuse (inheritance or embedding) | Proxy, Bridge, Composite, Chain of Responsibility, Template Method | struct embedding | Doesn't affect functionality — all patterns already work with interface + composition. But manual delegation creates lots of boilerplate; Go's embedding solves this directly |
| P1 | Method overloading / double dispatch | Visitor only | type switch | Narrowest scope, but Visitor genuinely needs double dispatch to be idiomatic |
| P2 | Access control modifiers | Singleton, Memento | First-letter case controls visibility | Small impact; Singleton's private constructor is the only hard need |

> **Already resolved**: `For Each` custom collection iteration (originally P0) was implemented by the author (issue #52); the Iterator pattern disappears entirely.

### Summary

| Status | Count |
|------|------|
| Classic already idiomatic | 3 |
| AxonASP fully resolves | 18 |
| AxonASP improves but residual gaps | 2 |
| **Idiomatic total** | **21** |

AxonASP raises idiomatic implementations from **3/23** to **21/23**. The 18 "fully resolved" patterns use `Implements` interfaces, `Static` variables, `Event` mechanisms, `For Each` custom collection iteration, and other extensions to achieve compile-time type-safe polymorphic dispatch and native traversal — no helper classes or fully-qualified names needed.

The residual gaps in the remaining 2 patterns fall into two categories when compared to Go:

- **Go also lacks it but has an alternative**: Method overloading / double dispatch (Go uses type switch). AxonASP currently uses interface + manual branching — functionally equivalent but more boilerplate.
- **Go has it but AxonASP doesn't**: Access control (Go's first-letter case). Singleton's private constructor is the only hard need.

Go lacks "inheritance" and "method overloading" — two features widely considered "essential for OOP" — yet can still idiomatically implement all 23 GoF patterns. This proves these features aren't required for design patterns — interface + composition is the core. AxonASP already has interfaces and composition. The direction is right.

### V. 8 AxonASP Features Used in This Book

AxonASP introduces many modern features to VBScript; this book uses 8 of them. Filtered by "does it help make the 23 patterns more idiomatic", 4 are core and 4 are nice-to-have.

**Core (passed the GoF-23 filter):**

| Feature | Used in | Why it's core |
|---|---|---|
| `Implements` interface | 17 patterns' polymorphic dispatch | Without it, Factory, Proxy, Adapter, Strategy etc. can only rely on method-name conventions — not idiomatic |
| `Event`/`RaiseEvent`/`WithEvents` | Ch.14 Observer | Without it, Observer requires manual array maintenance — not idiomatic |
| `Static` static variable | Ch.1 Singleton | Without it, Singleton requires a global variable — breaks encapsulation |
| `For Each` custom collection + `[DispId(-4)]` | Ch.16 Iterator | Without it, Iterator needs hand-written `HasNext`/`NextItem` classes; with it, the Iterator pattern disappears entirely |

**Nice-to-have (didn't pass the GoF-23 filter):**

| Feature | Used in | Why it doesn't pass the filter | Value |
|---|---|---|---|
| Strong typing (`As Type`) | 17 "fully resolved" patterns (except #16 Iterator due to built-in Collection COM object) + 3 "classic idiomatic" patterns' subsystem refs, fields, and params — 20 patterns total | All 23 patterns work fine with Variant; type safety isn't an idiomatic pattern issue | Eliminates entire classes of runtime type errors; foundation for IDE intellisense |
| UDT (`Type`) | Ch.22 Memento (`EditorMemento`) | Memento is already idiomatic with `Class` + `Property`; UDT is just cleaner | Value semantics naturally prevent reference leaks; less boilerplate |
| `Enum` enumeration | Ch.17 Chain of Responsibility (`LogLevel`) | Could use `Const` instead; functionally equivalent | Better readability; IDE auto-complete support |
| `IsNot` operator | Ch.17 Chain of Responsibility (`m_Next IsNot Nothing`) | `Not x Is Nothing` is functionally equivalent; `IsNot` is just clearer | Eliminates `Not ... Is Nothing` precedence ambiguity; guard clauses are more readable |

Without the 4 core features, 18 of the 23 patterns can't be idiomatically implemented. Without the 4 nice-to-have features, all 23 patterns are still idiomatic — but the code is safer, cleaner, and easier to maintain. Strong typing already covers 17 "fully resolved" patterns (Iterator #16 can't fully annotate due to COM interop limits) and 3 "classic idiomatic" patterns — 20 total; only Ch.16 Iterator retains Variant due to COM interop. This shows "nice-to-have" doesn't mean "insignificant" — it makes already-idiomatic code safer and more maintainable. If AxonASP introduces more general language improvements in the future (like `AndAlso`/`OrElse` short-circuit evaluation, parameterized constructors, etc.), the GoF-23 filter can similarly be used to evaluate their priority — consistent with the classification here.

---

## Quick Reference: 23 Patterns at a Glance

| # | Pattern | Category | One-line summary |
|:---:|---|---|---|
| 1 | **Singleton** | Creational | Only one instance globally |
| 2 | **Factory Method** | Creational | Delegate creation decision to subclass/factory |
| 3 | **Abstract Factory** | Creational | Swap factory, swap whole product family |
| 4 | **Builder** | Creational | Build complex objects step by step |
| 5 | **Prototype** | Creational | Create new objects by copying |
| 6 | **Proxy** | Structural | Provide a surrogate / lazy loading |
| 7 | **Facade** | Structural | Simple entry point for a complex system |
| 8 | **Adapter** | Structural | Convert incompatible interface to target |
| 9 | **Bridge** | Structural | Separate abstraction from implementation |
| 10 | **Composite** | Structural | Treat individuals and compositions uniformly |
| 11 | **Decorator** | Structural | Dynamically add behavior without modifying class |
| 12 | **Flyweight** | Structural | Share fine-grained objects to save memory |
| 13 | **Strategy** | Behavioral | Encapsulate algorithm as swappable object |
| 14 | **Observer** | Behavioral | Notify all watchers on state change |
| 15 | **Template Method** | Behavioral | Define algorithm skeleton; subclasses fill in steps |
| 16 | **Iterator** | Behavioral | Sequential access without exposing internals |
| 17 | **Chain of Responsibility** | Behavioral | Pass request along chain until handled |
| 18 | **Command** | Behavioral | Encapsulate request as object |
| 19 | **State** | Behavioral | Behavior changes with internal state |
| 20 | **Mediator** | Behavioral | Use mediator to encapsulate object interactions |
| 21 | **Visitor** | Behavioral | Encapsulate operations in a visitor |
| 22 | **Memento** | Behavioral | Save and restore object state |
| 23 | **Interpreter** | Behavioral | Build an interpreter for a language |
