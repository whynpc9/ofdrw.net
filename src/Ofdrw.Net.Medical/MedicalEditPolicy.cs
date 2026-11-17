using System;

namespace Ofdrw.Net.Medical;

/// <summary>Document operations distinguished by GB/T 48666 9.2(b), 7.3(g) and 6.3.</summary>
public enum MedicalEditOperation
{
    /// <summary>Insert a page.</summary>
    InsertPage,
    /// <summary>Remove a page.</summary>
    RemovePage,
    /// <summary>Reorder or rotate pages.</summary>
    AdjustPages,
    /// <summary>Change existing page content, including text replacement.</summary>
    EditPageContent,
    /// <summary>Add an annotation (6.3(f), 7.3(g)).</summary>
    AddAnnotation,
    /// <summary>Add or change description metadata (7.3(g)).</summary>
    EditMetadata,
    /// <summary>Add a signature or seal (7.3(g)).</summary>
    Sign,
    /// <summary>Combine documents (3.10, 7.3(g)).</summary>
    Combine
}

/// <summary>Decision for one operation on one declared document type.</summary>
public enum MedicalEditDecision
{
    /// <summary>The operation is not restricted by GB/T 48666.</summary>
    Allowed,
    /// <summary>Allowed, but existing signatures stop covering the result: remove and re-sign, or add a signature (3.10, 6.3(g)).</summary>
    AllowedWithResign,
    /// <summary>A reader disables it for OFD-H by default; only an authorized workflow may perform it (9.2(b)).</summary>
    RequiresAuthorization
}

/// <summary>Thrown when an OFD-H page-editing operation is attempted without authorization.</summary>
public sealed class MedicalEditNotAuthorizedException : InvalidOperationException
{
    internal MedicalEditNotAuthorizedException(MedicalEditOperation operation)
        : base("GB/T 48666 9.2(b): " + operation + " is disabled for DocType=OFD-H unless the caller's workflow authorizes it.")
    { Operation = operation; }
    /// <summary>The refused operation.</summary>
    public MedicalEditOperation Operation { get; }
}

/// <summary>
/// Classifies edit operations for OFD-H inputs. Ofdrw.Net's general editing APIs do not call this policy; applications
/// that edit medical records call it before editing. It never inspects signatures or modifies documents.
/// </summary>
public static class MedicalEditPolicy
{
    /// <summary>Classifies an operation for a raw declared <c>DocType</c>.</summary>
    /// <param name="declaredDocType">The raw root attribute value, or null.</param>
    /// <param name="operation">The intended operation.</param>
    /// <returns>Allowed for documents that are not exactly OFD-H; otherwise the GB/T 48666 decision.</returns>
    public static MedicalEditDecision Classify(string? declaredDocType, MedicalEditOperation operation)
    {
        if (!MedicalDocTypes.IsMedical(declaredDocType)) return MedicalEditDecision.Allowed;
        switch (operation)
        {
            case MedicalEditOperation.InsertPage:
            case MedicalEditOperation.RemovePage:
            case MedicalEditOperation.AdjustPages:
            case MedicalEditOperation.EditPageContent:
                return MedicalEditDecision.RequiresAuthorization;
            case MedicalEditOperation.EditMetadata:
            case MedicalEditOperation.Combine:
                return MedicalEditDecision.AllowedWithResign;
            case MedicalEditOperation.AddAnnotation:
            case MedicalEditOperation.Sign:
                return MedicalEditDecision.Allowed;
            default:
                throw new ArgumentOutOfRangeException(nameof(operation));
        }
    }

    /// <summary>Throws when the operation requires authorization that the caller has not granted.</summary>
    /// <param name="declaredDocType">The raw root attribute value, or null.</param>
    /// <param name="operation">The intended operation.</param>
    /// <param name="authorized">Whether the caller's workflow authorized page editing.</param>
    /// <exception cref="MedicalEditNotAuthorizedException">OFD-H page editing without authorization.</exception>
    public static void EnsureAllowed(string? declaredDocType, MedicalEditOperation operation, bool authorized)
    {
        if (Classify(declaredDocType, operation) == MedicalEditDecision.RequiresAuthorization && !authorized)
            throw new MedicalEditNotAuthorizedException(operation);
    }
}
