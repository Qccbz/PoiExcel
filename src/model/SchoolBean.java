package model;

public class SchoolBean {
	String name;
	int startPos;
	int endPos;

	public SchoolBean(String name, int startPos, int endPos) {
		super();
		this.name = name;
		this.startPos = startPos;
		this.endPos = endPos;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getStartPos() {
		return startPos;
	}

	public void setStartPos(int startPos) {
		this.startPos = startPos;
	}

	public int getEndPos() {
		return endPos;
	}

	public void setEndPos(int endPos) {
		this.endPos = endPos;
	}

}