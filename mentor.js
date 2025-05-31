function getTraineesDepartmentChoices(mentor) {
    const className = mentor['班級'];
    const [headers, ...data] = studentChoiceSheet.getDataRange().getValues();

    const classIndex = headers.indexOf('班級名稱');
    const serialIndex = headers.indexOf('統一入學測驗報名序號');
    const studentNameIndex = headers.indexOf('考生姓名');
    const groupIndex = headers.indexOf('報考群(類)名稱');
    const paymentTypeIndex = headers.indexOf('繳費身分');
    const feeIndex = headers.indexOf('報名費');
    const isJoinedIndex = headers.indexOf('是否參加集體報名');
    const choiceIndex1 = headers.indexOf('志願1校系名稱');
    const choiceIndex2 = headers.indexOf('志願2校系名稱');
    const choiceIndex3 = headers.indexOf('志願3校系名稱');
    const choiceIndex4 = headers.indexOf('志願4校系名稱');
    const choiceIndex5 = headers.indexOf('志願5校系名稱');
    const choiceIndex6 = headers.indexOf('志願6校系名稱');

    // 篩選出該班級的學生資料
    const mentorStudents = data.filter((row) => row[classIndex] === className);

    return {
        headers: [
            '考生姓名',
            '統一入學測驗報名序號',
            '班級名稱',
            '報考群(類)名稱',
            '繳費身分',
            '報名費',
            '是否參加集體報名',
            '志願1校系名稱',
            '志願2校系名稱',
            '志願3校系名稱',
            '志願4校系名稱',
            '志願5校系名稱',
            '志願6校系名稱',
        ],
        data: mentorStudents.map((row) => [
            row[studentNameIndex],
            row[serialIndex],
            row[classIndex],
            row[groupIndex],
            row[paymentTypeIndex],
            row[feeIndex],
            row[isJoinedIndex],
            row[choiceIndex1],
            row[choiceIndex2],
            row[choiceIndex3],
            row[choiceIndex4],
            row[choiceIndex5],
            row[choiceIndex6],
        ]),
    };
}
