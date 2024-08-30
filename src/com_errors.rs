
pub enum ComError{
    NotInitialize(),
    ComInstance(),
    NoInterface(String),
    NoMethod(String),
    ComNotFound(),
    PointerAlreadyMapped()
}