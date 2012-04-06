CREATE OR REPLACE PROCEDURE procPopularCourse(cName OUT varchar2)
AS
	cID number(4) := 0 ;
	cNum number(4) := 0 ;
	cRec cm_course%ROWTYPE ;

BEGIN

	SELECT MAX(COUNT(student_id)) INTO cNum 
	FROM cm_student_course GROUP BY course_id ;

	SELECT "course" INTO cID FROM
	(
		SELECT COUNT(student_id) "num", course_id "course"
		FROM cm_student_course
		GROUP BY course_id
		ORDER BY "num" DESC
	)
	WHERE rownum = 1 ;

	SELECT * INTO cRec FROM cm_course
	WHERE course_id = cID ;

	cName := cRec.class || 'th ' || cRec.course_name || ' ' || cRec.type ;

END procPopularCourse;