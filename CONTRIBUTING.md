# Contributing to EditControls

The following is a set of guidelines for contributing to the EditControls project. These are mostly guidelines, not rules. Use your best judgment, and feel free to propose changes to this document in a pull request.

#### Table Of Contents
[Styleguides](#styleguides)
  * [Git Commit Messages](#git-commit-messages)
  * [C++ Styleguide](#c-styleguide)

## Styleguides

### Git Commit Messages

* Write them in English. It does not have to be perfect.
* Ask yourself whether after 6 months you or anybody else will still be able to understand what has been changed and why it has been changed.
* If external information (e.g. documentation at MSDN or threads at StackOverflow) has been important to your commit, link it.

### C++ Styleguide

In most cases we follow the style suggested by the [ECMA-334](http://www.ecma-international.org/publications/files/ECMA-ST/Ecma-334.pdf) standard.

(1) Terminology:
  - raising an event means preparing its sending to the connected clients
  - firing an event means sending it to all connected clients

(2) Names
  - macros and constants: Upper Case (```MACRO```)
  - functions, typedefs, structs: Pascal Case (```MyMethod```)
  - variables: Camel Case (```myVariable```)
  - use self-explanatory names

(3) Prefixes
  - interfaces - I
  - HANDLE - h
  - pointer - p
  - pointer to a function - pfn
  - No Hungarian notation! (except above exceptions)

(4) Indentation
  - mulit-line instructions: 4 white spaces
  - trailing comments: 5 white spaces
  - else: 1 tab
  - brackets don't get indented

(5) Comments
  - multi-line comments: ```/* */```
        ```/* This is a sample.
           The second line is indented using spaces. */```
  - single-line comments: ```//```
  - trailing comments: ```//```
  - inline comments: ```/* */```
  - consider commenting closing scope brackets ```}``` if the corresponding open one is located many lines before
  - TODO for optional improvements that may be useful
  - HACK for ugly code
  - UNDONE for incomplete implementations
  - FIXME for bugs
  - NOTE for simple, but important notes

(5a) Doxygen
  - comment style: ```///```
  - class docs are placed at the top of the file, surrounded by ```/////////////////////////``` lines
  - simple text formatings are done using html tags

(6) Blank lines
  - use them to separate logical blocks (which should also have an introducing block comment)

(7) Brackets
  - scope brackets ```{}``` start on the same line as the corresponding order (separated by a single white space) and end on a separate line
       - exceptions: namespaces, classes, interfaces, methods and structs (where the opening bracket is located on a new line)
  - logical brackets ```()``` don't get surrounded by white spaces
  - always use ```{}```, also if they're not neccessary:
       ```
       if(myBool) {
         return S_OK;
       }
       ```

(8) Pointers
  - ```int* pMyInt```, NOT ```int *pMyInt```
  - ```int& myInt```, NOT ```int &myInt```

(9) Macros, inline methods, register variables, asm etc.
  - use it where speed matters and/or it significantly shortens the code

(10) Friend classes
  - avoid it if possible

(11) Namespaces
  - prefer STL for vectors, lists and similar
  - prefer ATL for strings and COM-related stuff
  - prefer WTL for GUI and things like CPoint
  - ATL is the only global namespace



#### Templates

```
class MyClass :
    public MyBaseClass
{
public:
  long tmp;

  void MySub1(void);

protected:
  int MySub2(int i = 20);

private:
  bool flag;
}     // MyClass

for(int i = 0; i < 10; ++i) {
  a[i] = i;
}

if(b) {
  return S_OK;
} else {
  return E_FAIL;
}

if((myInt == 10) && (hr == S_OK)) {
  // do something
}

switch(myLong) {
  case 0:
    DoSomething();
    break;
  case 1:
    DoSomething();
    break;
  default:
    myLong = 10;
    break;
}     // switch(myLong)
```
