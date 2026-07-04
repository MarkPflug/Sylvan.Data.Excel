#if !NET5_0_OR_GREATER
// Polyfill so trim-annotation attributes can be applied when multi-targeting
// TFMs that predate their introduction in the BCL (.NET 5). These have no
// runtime or analyzer effect on those older targets; they exist only so the
// annotation on ExcelDataReader.GetFieldType compiles across all TFMs and
// matches DbDataReader.GetFieldType on the TFMs where it is actually enforced.
namespace System.Diagnostics.CodeAnalysis;

[Flags]
enum DynamicallyAccessedMemberTypes
{
	None = 0,
	PublicParameterlessConstructor = 0x0001,
	PublicConstructors = 0x0002 | PublicParameterlessConstructor,
	NonPublicConstructors = 0x0004,
	PublicMethods = 0x0008,
	NonPublicMethods = 0x0010,
	PublicFields = 0x0020,
	NonPublicFields = 0x0040,
	PublicNestedTypes = 0x0080,
	NonPublicNestedTypes = 0x0100,
	PublicProperties = 0x0200,
	NonPublicProperties = 0x0400,
	PublicEvents = 0x0800,
	NonPublicEvents = 0x1000,
	Interfaces = 0x2000,
	All = -1,
}

[AttributeUsage(
	AttributeTargets.Field | AttributeTargets.ReturnValue | AttributeTargets.GenericParameter |
	AttributeTargets.Parameter | AttributeTargets.Property,
	Inherited = false)]
sealed class DynamicallyAccessedMembersAttribute : Attribute
{
	public DynamicallyAccessedMembersAttribute(DynamicallyAccessedMemberTypes memberTypes) =>
		MemberTypes = memberTypes;

	public DynamicallyAccessedMemberTypes MemberTypes { get; }
}
#endif
