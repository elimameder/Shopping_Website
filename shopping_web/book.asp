
<html>
<head>
    <title>Shopping Website</title>
    <link rel="stylesheet" href="style.css">
    <link rel="stylesheet"
          href="https://cdn.jsdelivr.net/npm/boxicons@latest/css/boxicons.min.css">
</head>
<body>
    <header>
        <a href="#" class=" logo"> Brand <span>X.</span></a>
        <div class="bx bx-menu" id="menu-icon"></div>
        <ul class=" navbar">
            <li><a href="#home">Home</a></li>
            <li><a href="#shop">Shop</a></li>
            <li><a href="#new">New Arrival</a></li>
            <li><a href="#brands">Brands</a></li>
            <li><a href="#about">About</a></li>
            <li><a href="#contact">Contacts</a></li>
        </ul>
    </header>
    <section class=" home" id=" home">
        <div class=" home-text">
            <h1><Span>Turn</Span>Yourself<br> Into <span>a Brand</span></h1>
            <p>LOREMMMMMMMMMMMMMMMMMMMMMM</p>
            <a href=" shop" class=" btn">Shop Now</a>
        </div>
    </section>
    <section class=" shop" id=" shop">
        <div class=" heading">
            <span>New Arrival</span>
            <h2>Shop Now</h2>
        </div>

        <div class="shop-container">

            <div class=" box">
                <div class="box-img">
                    <img src="dress2.jpg" alt="">
                </div>
                <div class=" title-price">
                    <h3>Silver Dress </h3>
                    <div class=" stars">
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star-half'></i>
                    </div>
                </div>
                <span>98$</span>
                <i class='bx bxs-cart'></i>
            </div>

            <div class=" box">
                <div class="box-img">
                    <img src="dress4.png" alt="">
                </div>
                <div class=" title-price">
                    <h3>Silver Dress </h3>
                    <div class=" stars">
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star-half'></i>
                    </div>
                </div>
                <span>98$</span>
                <i class='bx bxs-cart'></i>
            </div>

            <div class=" box">
                <div class="box-img">
                    <img src="dress5.png" alt="">
                </div>
                <div class=" title-price">
                    <h3>Silver Dress </h3>
                    <div class=" stars">
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star-half'></i>
                    </div>
                </div>
                <span>98$</span>
                <i class='bx bxs-cart'></i>
            </div>

            <div class=" box">
                <div class="box-img">
                    <img src="d3.png" alt="">
                </div>
                <div class=" title-price">
                    <h3>Silver Dress </h3>
                    <div class=" stars">
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star-half'></i>
                    </div>
                </div>
                <span>98$</span>
                <i class='bx bxs-cart'></i>
            </div>

            <div class=" box">
                <div class="box-img">
                    <img src="d2 (3).png" alt="">
                </div>
                <div class=" title-price">
                    <h3>Silver Dress </h3>
                    <div class=" stars">
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star-half'></i>
                    </div>
                </div>
                <span>98$</span>
                <i class='bx bxs-cart'></i>
            </div>

            <div class=" box">
                <div class="box-img">
                    <img src="dress6.png" alt="">
                </div>
                <div class=" title-price">
                    <h3>Silver Dress </h3>
                    <div class=" stars">
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star'></i>
                        <i class='bx bxs-star-half'></i>
                    </div>
                </div>
                <span>98$</span>
                <i class='bx bxs-cart'></i>
            </div>
        </div>
    </section>
    <section class="contact" id=" contact">
        <h1 class="heading"><span>Contact</span> Us</h1>

<%
Dim adoCon 			
Dim rsGuestbook		
Dim strSQL			

Set adoCon = Server.CreateObject("ADODB.Connection")


adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("DB.mdb")


Set rsGuestbook = Server.CreateObject("ADODB.Recordset")

strSQL = "SELECT Table1.Name, Table1.Email, Table1.Phone, Table1.Subject FROM Table1;"

rsGuestbook.Open strSQL, adoCon

Do While not rsGuestbook.EOF
	

	Response.Write ("<br>")
	Response.Write (rsGuestbook("Name"))
	Response.Write ("<br>")
	Response.Write (rsGuestbook("Email"))
	Response.Write ("<br>")
    Response.Write (rsGuestbook("Phone"))
	Response.Write ("<br>")
    Response.Write (rsGuestbook("Subject"))
	Response.Write ("<br>")


	rsGuestbook.MoveNext

Loop

rsGuestbook.Close
Set rsGuestbook = Nothing
Set adoCon = Nothing
%>
    </section>

</body>
</html>