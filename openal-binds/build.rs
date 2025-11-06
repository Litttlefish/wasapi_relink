use std::env;
use std::path::PathBuf;

fn main() {
    println!("cargo:rerun-if-changed=wrapper.h");
    let lib_dir = if cfg!(target_pointer_width = "64") {
        // 64位目标
        concat!(env!("CARGO_MANIFEST_DIR"), "/libs/Win64")
    } else {
        // 32位目标 (或其他非64位架构)
        concat!(env!("CARGO_MANIFEST_DIR"), "/libs/Win32")
    };
    println!("cargo:rustc-link-search=native={}", lib_dir);
    println!("cargo:rustc-link-lib=OpenAL32");

    let bindings = bindgen::Builder::default()
        .header("wrapper.h")
        .parse_callbacks(Box::new(bindgen::CargoCallbacks::new()))
        .generate()
        .expect("Unable to generate bindings");

    let out_path = PathBuf::from(env::var("OUT_DIR").unwrap());
    bindings
        .write_to_file(out_path.join("bindings.rs"))
        .expect("Couldn't write bindings!");
}
