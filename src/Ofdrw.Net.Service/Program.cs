using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.EmrTechSpec.Abstractions;
using Ofdrw.Net.EmrTechSpec.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddOpenApi();

builder.Services.AddSingleton<IPdfToOfdConverter, PdfToOfdConverter>();
builder.Services.AddSingleton<IOfdToPdfConverter, OfdToPdfConverter>();
builder.Services.AddSingleton<IEmrTechSpecValidator, EmrTechSpecValidator>();
builder.Services.AddSingleton<EmrValidationProfileRepository>();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

app.UseHttpsRedirection();
app.MapControllers();

app.Run();

public partial class Program;
