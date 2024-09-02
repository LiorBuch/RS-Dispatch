use std::fmt::Display;


pub enum ComError{
    NotInitialize(),
    ComInstance(String),
    NoInterface(String),
    NoMethod(String),
    ComNotFound(),
    PointerAlreadyMapped()
}

impl Display for ComError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ComError::NotInitialize() => write!(f,"Com initialize error."),
            ComError::ComInstance(st) => write!(f,"Com instance creation error: {st}"),
            ComError::NoInterface(_) => write!(f,"No interface with that name was found."),
            ComError::NoMethod(_) => write!(f,"No method with that name was found."),
            ComError::ComNotFound() => write!(f,"Com object was not found."),
            ComError::PointerAlreadyMapped() => write!(f,"Com pointer already mapped."),
        }
    }
}