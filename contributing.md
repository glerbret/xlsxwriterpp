# Reporting Bugs and submitting Pull Requests

## Reporting Bugs

### Upgrade to the latest version of the library

Upgrade to the latest version of the library since the bug you are reporting
may already be fixed.

Check the changelog.txt to see what has changed in the latest versions.

You can check which version of `Xlsxwriter++` that you are using by checking
the `xlsxwriterpp.h` header file or by adding the following to your program:

```cpp
#include "xlsxwriterpp.h"

#include <format>
#include <iostream>

int main()
{
    std::cout << std::format("Xlsxwriter++ version: {}\n", xwpp::version());
}
```

### Read the documentation

Read or search the `libxlsxwriter` documentation to see if the issue
you are encountering is already explained.

### Look at the example programs

There are many examples programs in the distribution. Try to
identify an example program that corresponds to your query and adapt it to use
as a bug report.

### Use the Xlsxwriter++ Issue Tracker

The [Xlsxwriter++ issue tracker][issues] is on GitHub.

[issues]: https://github.com/glerbret/xlsxwriterpp/issues

### Tips for submitting a bug report

1. Describe the problem as clearly and as concisely as possible.
1. Include a sample program. This is probably the most important step.
   It is generally easier to describe a problem in code than in written
   prose.
1. The sample program should be as small as possible to demonstrate the
   problem. Don't copy and paste large non-relevant sections of your
   program.

A sample bug report is shown below. This format helps analyze and respond to
the bug report more quickly.

> Subject: Issue with SOMETHING
>
> Greetings,
>
> I am using Xlsxwriter++ to do SOMETHING but it appears to do SOMETHING ELSE.
>
> I am using CXX version X.Y.Z, OS = uname and Xlsxwriter++ x.y.z.
>
> Here is some code that demonstrates the problem:
>
>
>```cpp
>#include "xlsxwriterpp.h"
>
>int main()
>{
>    xwpp::workbook_t workbook;
>    xwpp::worksheet_t& worksheet = workbook.add_worksheet();
>
>    worksheet.write(0, 0, "Hello");
>    worksheet.write(1, 0, 123);
>
>     workbook.save("bug_report.xlsx");
>}
>```
>

## Pull Requests and Contributing to Xlsxwriter++

All patches and pull requests are welcome but in general you should start with
an issue tracker to describe what you intend to do before you do it.

### Getting Started

1. Pull requests and new feature proposals must start with an [issue
   tracker][issues]. This serves as the focal point for the design discussion.
2. Describe what you plan to do. If there are API changes add some code
   example to demonstrate them.
3. Fork the repository.
4. Run all the tests to make sure the current code works on your system.
5. Create a feature branch for your new feature.

### Code Style

The code style is mainly with 2 space indents.

The [clang-format configuration file][clangformat] is available in the repo. The code can be indented automatically with the `format.sh` command.

[clangformat]: https://github.com/glerbret/xlsxwriterpp/blob/main/.clang-format

### Naming convention

* All names use snake_case style.
* Function names must describe action.
* Type names (class, struct, enum class, ...) have a `_t` suffix.
* Private data members have a `_` suffix.

### Writing and Running Tests

Any significant features should be accompanied by a test. See the `test`
directory.

The functional tests require the Python module [pytest][pytest] as a test runner.

When you push your changes they will also be tested automatically using
[GitHub Actions][actions].

[actions]: https://github.com/glerbret/xlsxwriterpp/actions
[pytest]: http://pytest.org/

### Documentation

The `Xlsxwriter++` documentation is written in Doxygen format in the header
files and in additional `.dox` files in the `docs/src` directory of the
repo. The documentation can be built with `Doxygen`.

### Copyright and License

Copyright remains with the original author. Do not include additional
copyright claims or Licensing requirements. GitHub and the `git` repository
will record your contribution and it will be acknowledged it in the Changes
file.

### Submitting the Pull Request

If your change involves several incremental `git` commits then `rebase` or
`squash` them onto another branch so that the Pull Request is a single commit
or a small number of logical commits.

Push your changes to GitHub and submit the Pull Request with a hash link to
the to the Issue tracker that was opened above.
