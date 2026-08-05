# 23 Design Patterns Explained — VBScript Edition

## About This Book

This book systematically covers the 23 classic design patterns from the GoF (Gang of Four) book *Design Patterns: Elements of Reusable Object-Oriented Software*, using **VBScript** as the vehicle. It compares implementations across two runtimes: **classic VBScript** and **AxonASP**. Each pattern includes runnable examples, with trade-offs and missing language features noted for both runtimes, so readers can see firsthand how language features affect pattern implementation.

## Why VBScript for Design Patterns

VBScript is a weakly-typed, COM-based scripting language born in 1996, still running in Classic ASP, Windows Script Host, embedded automation, and many legacy systems. Its syntax is minimal — just `Class`/`Function`/`Property`/`Dim` — with no inheritance, no interfaces, no generics, no static variables, no event mechanism. This makes it a perfect "anti-textbook": implementing 23 GoF patterns in a language with the fewest features magnifies every compromise and workaround.

In other words, **VBScript isn't the best language for showing how design patterns should be implemented — it's the best language for understanding why design patterns are designed the way they are**. Once you see "why Factory Method feels awkward without interfaces", "why Template Method gets distorted without inheritance", and "why Observer is bloated without events", going back to Java/C#/Go implementations takes on a whole new meaning.

## Two Runtimes

This book compares two VBScript runtimes:

- **Classic VBScript**: Standard VBScript environments like WSH and Classic ASP. Only basic syntax: `Class`, `Function`, `Property`, `Dim`. No inheritance, interfaces, polymorphism, overloading, static variables, events, or generics.
- **AxonASP**: A runtime that extends classic VBScript with modern features: `Implements` interface polymorphism, `Static` static variables, `Event`/`RaiseEvent`/`WithEvents` event mechanism, `For Each` custom collection iteration, strong typing (`As Type`), UDTs (`Type`), `Enum` enumerations, `IsNot` operator, and more. These extensions let most GoF patterns be implemented in a more idiomatic way.

Both versions' code can run independently: the classic version targets Classic ASP / WSH; the AxonASP version targets the AxonASP runtime. All examples output via `Response.Write` for consistent observation in ASP contexts.

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
5. If AxonASP still has unresolved gaps, they're explicitly marked as **missing syntax**, with comparisons to Go's equivalent approach.

The appendix provides an implementation status overview for all 23 patterns, missing syntax priorities, and the 8 AxonASP features actually used in this book.

## How to Read

- **Beginners**: Read in chapter order. Start with "Core idea" and "Example", then compare both code versions and read the trade-off notes.
- **Readers with OOP background**: Jump straight to the Axon VBScript versions. Focus on the "missing syntax" sections to understand the gap between AxonASP and Java/C#/Go.
- **VBScript veterans**: Focus on the classic version trade-offs — these are the pain points you've hit over 20 years of VBScript coding. This book names each one and shows the solution.
- **AxonASP users**: All Axon versions in this book have been tested against the AxonASP runtime and can serve as reference templates for pattern implementation.

## Conventions

- All `Response.Write` output in code is unbuffered single-line output; multi-line displays are for readability only.
- Comments are in Chinese, matching the original code language.
- **Missing syntax** items in trade-off notes are always bolded and correspond to the priority table in the appendix.
- Go language comparisons follow Go 1.18+ syntax (including generics); where AxonASP's current state is equivalent to pre-Go 1.18, this is explicitly noted.

## Acknowledgments

The AxonASP examples in this book benefit from the ongoing evolution of the AxonASP runtime. In particular, `For Each` custom collection iteration (issue #52) makes the Iterator pattern in Chapter 16 "disappear entirely", and the `New ClassName` inline function argument compilation bug fix simplified how concrete factories are written across multiple patterns. Thanks to all community contributors who contributed code and filed issues for AxonASP.

---
