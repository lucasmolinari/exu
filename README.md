# EXU
This project is a reimplementation of a old Python Script that I made, designed to unlock protected Excel sheets.

## Installation

1. **Clone the Repository**
    ```sh
    git clone https://github.com/lucasmolinari/exu.git
    cd exu
    ```

2. **(Optional) Install Rust**
   
    If you don't have Rust installed, follow the installation steps [here](https://www.rust-lang.org/tools/install).

3. **Build the Project**
    ```sh
    cargo build --release
    ```

## Usage

Run the executable with the path to the protected Excel file and the destination:
```sh
./target/release/exu path/to/protected.xlsx  path/to/destination/
```
or:
```sh
cargo run --release -- path/to/protected.xlsx path/to/destination/
```



