create table `test`.training_certificate
(
    id_card_number      varchar(25)  null comment '身份证号',
    certificate_type    varchar(25)  null comment '证书类型',
    certificate_name    varchar(25)  null comment '证书名称',
    certificate_company varchar(25)  null comment '发证单位',
    validity_enddate    varchar(25)     null comment '有效期截止时间',
    certificate_url     varchar(255) null comment '电子证书链接',
    remark              varchar(25)  null comment '备注'
)
    comment '培训证书列表';

create table `test`.training_record
(
    id_card_number   varchar(25)   null comment '身份证号',
    training_type    varchar(25)   null comment '培训类别',
    training_name    varchar(100)  null comment '培训名称',
    training_content varchar(1000) null comment '培训内容',
    training_time    varchar(100)   null comment '培训时间',
    training_company varchar(25)   null comment '培训单位',
    training_place   varchar(25)   null,
    training_teacher varchar(25)   null,
    training_hour    varchar(25)        null comment '培训课时',
    training_score   varchar(25)   null comment '成绩',
    source           varchar(25)   not null comment '数据来源',
    record_url       varchar(255)  null comment '学习记录详情链接'
)
    comment '培训档案';