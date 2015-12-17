# Mono CSharp Compiler
CSC=mcs
CSCFLAGS=/nologo /warn:4 /debug:pdbonly /o /nowarn:3003

# Resource compiler flags
#Fix resgen paths in mono
RESGEN=export MONO_IOMAP=all; resgen
RC=rc
RCFLAGS=/usesourcepath

# Source dirs
SRCDIR=./Trunk
PROPDIR=$(SRCDIR)/Properties
COLLDIR=$(SRCDIR)/Collections
UTILDIR=$(SRCDIR)/Utilities
PRIMDIR=$(SRCDIR)/Primitives

# Build dirs
BUILDDIR=$(SRCDIR)/build
APIBLDDIR=$(BUILDDIR)/HtmlToOpenXML

OPENXMLDLLDOC=index.xml
HTML2OPENXMLDLL=HtmlToOpenXML.dll

# Output
FILE_DLL=$(APIBLDDIR)/$(HTML2OPENXMLDLL)

# Documentation File
DOC=$(APIBLDDIR)/$(OPENXMLDLLDOC)

# Target
TARGET=/target:library

# References
REFERENCES=-r:OpenXMLLib.dll,System.Drawing,WindowsBase

# Sources
FILES_SRC= \
	$(SRCDIR)/HtmlDocumentStyle.cs \
	$(PROPDIR)/PredefinedStyles.Designer.cs \
	$(PROPDIR)/AssemblyInfo.cs \
	$(SRCDIR)/HtmlConverter.ProcessTag.cs \
	$(SRCDIR)/WebProxy.cs \
	$(SRCDIR)/StyleEventArgs.cs \
	$(SRCDIR)/HtmlEnumerator.cs \
	$(SRCDIR)/HtmlConverter.cs \
	$(SRCDIR)/ProvisionImageEventArgs.cs \
	$(COLLDIR)/RunStyleCollection.cs \
	$(COLLDIR)/HtmlAttributeCollection.cs \
	$(COLLDIR)/HtmlTableSpanCollection.cs \
	$(COLLDIR)/OpenXmlStyleCollectionBase.cs \
	$(COLLDIR)/OpenXmlDocumentStyleCollection.cs \
	$(COLLDIR)/NumberingListStyleCollection.cs \
	$(COLLDIR)/ParagraphStyleCollection.cs \
	$(COLLDIR)/TableContext.cs \
	$(COLLDIR)/TableStyleCollection.cs \
	$(UTILDIR)/ImageHeader.cs \
	$(UTILDIR)/OpenXmlExtension.cs \
	$(UTILDIR)/WebClientEx.cs \
	$(UTILDIR)/ImageProvisioningProvider.cs \
	$(UTILDIR)/Logging.cs \
	$(UTILDIR)/ConverterUtility.cs \
	$(UTILDIR)/HttpUtility.cs \
	$(PRIMDIR)/Unit.cs \
	$(PRIMDIR)/HtmlFont.cs \
	$(PRIMDIR)/SideBorder.cs \
	$(PRIMDIR)/UnitMetric.cs \
	$(PRIMDIR)/FontWeight.cs \
	$(PRIMDIR)/FontStyle.cs \
	$(PRIMDIR)/Margin.cs \
	$(PRIMDIR)/FontVariant.cs \
	$(PRIMDIR)/HtmlImageInfo.cs \
	$(PRIMDIR)/DataUri.cs \
	$(PRIMDIR)/HtmlBorder.cs \
	$(PRIMDIR)/HtmlTableSpan.cs \
	$(SRCDIR)/Configuration\ enum.cs

PREDEFSTYLES=$(PROPDIR)/PredefinedStyles

RESOURCE=\
	/resource:$(PREDEFSTYLES).resources

# Rules
$(FILE_DLL): $(PREDEFSTYLES).resources
	mkdir -p $(APIBLDDIR)
	$(CSC) $(CSCFLAGS) /out:$@ $(TARGET) $(FILES_SRC) $(RESOURCE) $(OXLIB) $(REFERENCES) /doc:$(DOC)

$(PREDEFSTYLES).resources: $(PREDEFSTYLES).resx
	$(RESGEN) $(RCFLAGS) $(PREDEFSTYLES).resx $@

build: $(FILE_DLL)

.PHONY: clean
clean:
	rm -rf $(APIBLDDIR)
	rm -f $(PREDEFSTYLES).resources