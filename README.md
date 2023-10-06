# xlconv

python script for encoding, decoding and passthrough of excel files to/from a weird CSV like multi sheet format that works well in version control systems like git

## Use case

In your repo root, if you have a .gitconfig like:

```
[filter "xlcsv"]
    smudge = xlconv.py -d
    clean = xlconv.py -e
```

and a .gitattributes like:

```
*.xlsx filter=xlcsv
```

and then run something like

```shell
git config --local include.path ../.gitconfig
```

you'll be able to work in excel locally but check in CSV like files that are much nicer for version control.

## Usage:
```shell
xlconv [options]
```

## Options:
- `-e, --encode:`  Encode the excel source file to the target file as a multi sheet csv.
- `-d, --decode:` Decode the multi sheet csv source file to the target excel  file.
- `-p, --passthrough:` Copy the source file to the target file passing it through an intermediate encode/decode step, useful for cleaning excel files.
- `-s, --source:` Specify the source file path. If not specified, the input will be read from stdin.
- `-t, --target:` Specify the target file path. If not specified, the output will be written to stdout.

## Examples:
1. Encode a file:
    ```shell
    xlconv -e -s input.txt -t encoded.txt
    ```

2. Decode a file:
    ```shell
    xlconv -d -s encoded.txt -t output.txt
    ```

3. Use passthrough mode:
    ```shell
    xlconv -p -s input.txt -t output.txt
    ```

4. Read input from stdin and write output to stdout:
    ```shell
    xlconv -d
    xlconv -e
    xlconv -p
    ```

## Notes:
- The script requires Python to be installed on your system.
- If only the `-s` option is provided, the input will be read from the specified file and written to stdout.
- If only the `-t` option is provided, the input will be read from stdin and written to the specified file.
- If neither the `-s` nor the `-t` options are provided, both the input and output will be read from and written to stdin and stdout respectively.
- If both the `-s` and `-t` options are provided, the input will be read from the specified file and written to the specified file.
- If any provided file paths are invalid or the specified files cannot be accessed, an error message will be displayed.


