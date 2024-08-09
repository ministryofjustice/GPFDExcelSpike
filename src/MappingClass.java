
public class MappingClass {
String sourceName;
String mappedName;
String dateFormat;
String numberFormat;



public String getSourceName() {
	return sourceName;
}
public void setSourceName(String sourceName) {
	this.sourceName = sourceName;
}
public String getMappedName() {
	return mappedName;
}
public void setMappedName(String mappedName) {
	this.mappedName = mappedName;
}
public String getDateFormat() {
	return dateFormat;
}
public void setDateFormat(String dateFormat) {
	this.dateFormat = dateFormat;
}
public String getNumberFormat() {
	return numberFormat;
}
public void setNumberFormat(String numberFormat) {
	this.numberFormat = numberFormat;
}

@Override
public String toString() {
	return "MappingClass [sourceName=" + sourceName + ", mappedName=" + mappedName + ", dateFormat=" + dateFormat
			+ ", numberFormat=" + numberFormat + "]";
}

}

