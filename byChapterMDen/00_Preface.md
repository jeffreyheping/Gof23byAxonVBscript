# 23 Design Patterns Explained — VBScript Edition

## About This Book

This book systematically covers the 23 classic design patterns from the GoF (Gang of Four) book *Design Patterns: Elements of Reusable Object-Oriented Software*, using **VBScript** as the primary vehicle. Implementations are compared across two syntax variants — **classic VBScript** and **AxonASP** — with **VB.NET** serving as a syntactically complete baseline for reference. Each pattern includes runnable examples, with trade-offs and missing language features noted for each runtime, so readers can see firsthand how language features affect pattern implementation.

## Why VBScript for Design Patterns

VBScript is a weakly-typed, COM-based scripting language born in 1996, still running in Classic ASP, Windows Script Host, embedded automation, and many legacy systems. Its syntax is minimal — just `Class`/`Function`/`Property`/`Dim` — with no inheritance, no interfaces, no generics, no static variables, no event mechanism. This makes it a perfect "anti-textbook": implementing 23 GoF patterns in a language with the fewest features magnifies every compromise and workaround.

In other words, **VBScript isn't the best language for showing how design patterns should be implemented — it's the best language for understanding why design patterns are designed the way they are**. Once you see "why Factory Method feels awkward without interfaces", "why Template Method gets distorted without inheritance", and "why Observer is bloated without events", going back to Java/C#/Go implementations takes on a whole new meaning.

## Four Runtimes

This book covers four runtimes across three syntax systems. Classic VBScript code is tested on three engines; Axon VBScript code runs on AxonASP; VB.NET code runs on .NET.

### 1. Classic VBScript (cscript / WSH)

Microsoft's native VBScript engine, running on Windows Script Host (`cscript.exe`/`wscript.exe`) and Classic ASP (IIS). This is VBScript in its purest form — minimal syntax with only `Class`, `Function`, `Property`, `Dim` and a few other building blocks. All variables are `Variant` (weakly typed). No inheritance, interfaces, polymorphism, overloading, static variables, events, or generics. The classic VBScript code in this book is written against this engine.

### 2. AxonASP ([github.com/guimaraeslucas/axonasp](https://github.com/guimaraeslucas/axonasp))

AxonASP is a cross-platform Classic ASP runtime (Linux / Windows / macOS) written from scratch in Go, with a built-in single-pass compiler and stack-based bytecode VM. While fully compatible with classic VBScript syntax, AxonASP introduces several modern extensions:

- **`Implements` interfaces**: Define and implement interfaces, giving polymorphic dispatch compile-time contract enforcement.
- **`Static` static variables**: Function-level `Static` variables persist across calls, supporting object references.
- **`Event` / `RaiseEvent` / `WithEvents`**: Full event mechanism — Observer pattern can be implemented natively.
- **`For Each` custom collection iteration**: Via `[DispId(-4)]` forwarding the built-in Collection's enumerator — the Iterator pattern "disappears entirely".
- **Strong typing (`As Type`)**: Fields, parameters, and return values can be type-annotated, eliminating entire classes of runtime type errors.
- **UDT (`Type`) / `Enum` / `IsNot` operator**: Value-semantic structs, named constant sets, clearer null checks.

These extensions let most GoF patterns be implemented in a more idiomatic way. The Axon VBScript code in this book runs on this engine.

### 3. ASPPY ([github.com/PieterCooreman/ASPPY](https://github.com/PieterCooreman/ASPPY))

ASPPY is a Classic ASP / VBScript runtime implemented in Python, running cross-platform on Python 3.8+. It implements the full Classic ASP object model (`Request`/`Response`/`Session`/`Application`/`Server`) and a nearly complete VBScript built-in function library. ASPPY's syntax is **identical** to classic VBScript — no extensions added. It validates the portability of the same classic code on a different engine. ASPPY compiles ASP pages to Python bytecode with caching, auto-recompiling on file changes. The classic VBScript code in this book is also tested on ASPPY.

### 4. VB.NET (.NET CLR)

VB.NET is a full-featured OOP language on the .NET platform, with inheritance, interfaces, generics, delegates, events, `Shared` static members, `Private` constructors, `ReadOnly` fields, and all other modern language features. Each chapter in this book includes a VB.NET implementation as a **syntactically complete baseline** — every compromise you're forced to make in VBScript has a "correct answer" in VB.NET for reference. VB.NET code is compiled with `Option Strict On` + `Option Explicit On`, eliminating implicit type conversions.

### Runtime-to-Code Mapping

| Runtime | Syntax system | Code version |
|--------|---------|-----------|
| Classic VBScript (cscript) | Classic VBScript | Classic VBScript version |
| AxonASP | Axon-extended VBScript | Axon VBScript version |
| ASPPY | Classic VBScript | Classic VBScript version |
| VB.NET (.NET) | VB.NET | VB.NET version |

Classic VBScript code is tested on three engines (cscript, AxonASP classic mode, ASPPY) to verify portability; Axon VBScript code runs on AxonASP; VB.NET code runs on .NET.

## Book Structure

The book organizes the 23 patterns following GoF's traditional classification:

- **Creational patterns (Chapters 1–5)**: Singleton, Factory Method, Abstract Factory, Builder, Prototype. Focus on "how objects are created".
- **Structural patterns (Chapters 6–12)**: Proxy, Facade, Adapter, Bridge, Composite, Decorator, Flyweight. Focus on "how objects are composed".
- **Behavioral patterns (Chapters 13–23)**: Strategy, Observer, Template Method, Iterator, Chain of Responsibility, Command, State, Mediator, Visitor, Memento, Interpreter. Focus on "how objects collaborate".

Each chapter follows a consistent structure:

1. **Core idea** — one sentence capturing the pattern's essence.
2. **Example** — the simplest story that illustrates the problem the pattern solves.
3. **Classic VBScript version** — complete runnable code with trade-off notes.
4. **Axon VBScript version** — complete runnable code with trade-off notes.
5. **VB.NET version** — syntactically complete baseline, showing an implementation "that needs no compromises".
6. If AxonASP still has unresolved gaps, they're explicitly marked as **missing syntax**, with comparisons to Go's equivalent approach.

The appendix provides an implementation status overview for all 23 patterns, missing syntax priorities, and the AxonASP features actually used in this book.

## How to Read

- **Beginners**: Read in chapter order. Start with "Core idea" and "Example", then compare all three code versions and read the trade-off notes.
- **Readers with OOP background**: Jump straight to the Axon VBScript and VB.NET versions. Focus on the "missing syntax" sections to understand the gap between AxonASP and Java/C#/Go.
- **VBScript veterans**: Focus on the classic version trade-offs — these are the pain points you've hit over 20 years of VBScript coding. This book names each one and shows the solution.
- **AxonASP users**: All Axon versions in this book have been tested against the AxonASP runtime and can serve as reference templates for pattern implementation.

## Conventions

- Classic VBScript and Axon VBScript output via `Response.Write`; VB.NET output via `Console.WriteLine`.
- All `Response.Write` output is unbuffered single-line output; multi-line displays are for readability only.
- Comments are in Chinese, matching the original code language.
- **Missing syntax** items in trade-off notes are always bolded and correspond to the priority table in the appendix.
- Go language comparisons follow Go 1.18+ syntax (including generics); where AxonASP's current state is equivalent to pre-Go 1.18, this is explicitly noted.

---
