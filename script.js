/* ===========================
   Sales Management System
   style.css
=========================== */

*{
    margin:0;
    padding:0;
    box-sizing:border-box;
}

body{
    background:#f4f7fb;
    font-family:Segoe UI,Tahoma,Geneva,Verdana,sans-serif;
}

.navbar{
    height:65px;
}

.card{
    border:none;
    border-radius:12px;
    transition:.3s;
}

.card:hover{
    transform:translateY(-3px);
    box-shadow:0 10px 20px rgba(0,0,0,.15);
}

.card-header{
    font-weight:600;
    font-size:18px;
}

.list-group-item{
    padding:15px;
    font-size:16px;
}

.list-group-item.active{
    background:#0d6efd;
    border:none;
}

.list-group-item:hover{
    background:#eef5ff;
    cursor:pointer;
}

h3{
    font-weight:bold;
}

input[type=text]{
    height:45px;
}

input[type=file]{
    height:45px;
}

.btn{
    height:45px;
    font-weight:600;
    border-radius:8px;
}

.btn:hover{
    transform:scale(1.02);
}

.table{
    margin:0;
}

.table th{
    text-align:center;
    vertical-align:middle;
}

.table td{
    vertical-align:middle;
}

.table tbody tr:hover{
    background:#f8fbff;
}

#uploadStatus{
    font-size:17px;
}

.footer{
    text-align:center;
    color:#777;
    margin-top:30px;
}

.dashboard-card{
    color:#fff;
    border-radius:12px;
}

.sales-card{
    background:#0d6efd;
}

.customer-card{
    background:#198754;
}

.product-card{
    background:#fd7e14;
}

.qty-card{
    background:#6f42c1;
}

@media(max-width:992px){

.col-md-3{
    margin-bottom:15px;
}

}

@media(max-width:768px){

.navbar-brand{
    font-size:18px;
}

.card-header{
    font-size:16px;
}

h3{
    font-size:24px;
}

.btn{
    margin-top:10px;
}

}

@media(max-width:576px){

.container-fluid{
    padding:15px;
}

table{
    font-size:13px;
}

}
