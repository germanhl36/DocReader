/// A global actor that serialises all PDF rendering work off the main thread.
///
/// Use `@DocRenderActor` to annotate rendering functions so that they never
/// block the main thread, while still running on a dedicated serial executor.
@globalActor
public actor DocRenderActor {
    public static let shared = DocRenderActor()
}
