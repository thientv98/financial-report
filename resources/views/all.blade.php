<!DOCTYPE html>
<html lang="en">

<!-- Mirrored from colorlib.com/etc/lf/Login_v17/index.html by HTTrack Website Copier/3.x [XR&CO'2014], Tue, 23 Feb 2021 04:21:34 GMT -->
<!-- Added by HTTrack -->
<meta http-equiv="content-type" content="text/html;charset=UTF-8" /><!-- /Added by HTTrack -->

<head>
    <title>Financial Report</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <link rel="icon" type="image/png" href="/images/icons/favicon.ico" />

    <link rel="stylesheet" type="text/css" href="/vendor/bootstrap/css/bootstrap.min.css">

    <link rel="stylesheet" type="text/css" href="/fonts/font-awesome-4.7.0/css/font-awesome.min.css">

    <link rel="stylesheet" type="text/css" href="/fonts/Linearicons-Free-v1.0.0/icon-font.min.css">

    <link rel="stylesheet" type="text/css" href="/vendor/animate/animate.css">

    <link rel="stylesheet" type="text/css" href="/vendor/css-hamburgers/hamburgers.min.css">

    <link rel="stylesheet" type="text/css" href="/vendor/animsition/css/animsition.min.css">

    <link rel="stylesheet" type="text/css" href="/vendor/select2/select2.min.css">

    <link rel="stylesheet" type="text/css" href="/vendor/daterangepicker/daterangepicker.css">

    <link rel="stylesheet" type="text/css" href="/css/util.css">
    <link rel="stylesheet" type="text/css" href="/css/main.css">

</head>

<body>
    <div class="limiter">
        <div class="container-login100">
            <div class="wrap-login100">
                <form class="login100-form validate-form" action="export-all" method="get">
                    <span class="login100-form-title p-b-34">
                        Download Report
                    </span>

                    <div class="wrap-input100 rs1-wrap-input100 validate-input m-b-20" data-validate="Year From">
                        <input id="year_from" class="input100" type="number" min="1990" name="year_from" placeholder="Year From" value="2015">
                        <span class="focus-input100"></span>
                    </div>

                    <div class="wrap-input100 rs1-wrap-input100 validate-input m-b-20" data-validate="Year End">
                        <input id="year_to" class="input100" type="number" min="1990" name="year_to" placeholder="Year End" value="2020">
                        <span class="focus-input100"></span>
                    </div>

                    <div class="validate-input m-b-20" style="width: 100%">
                        <div class="form-check">
                            <input class="form-check-input ml-0" type="radio" name="horizontal" value="1" checked>
                            <label class="form-check-label" for="flexRadioDefault1">
                                Horizontal
                            </label>
                        </div>
                        <div class="form-check">
                            <input class="form-check-input ml-0" type="radio" name="horizontal" value="0" id="flexRadioDefault2">
                            <label class="form-check-label" for="flexRadioDefault2">
                                Vertical
                            </label>
                        </div>
                    </div>
                    {{-- <div class="wrap-input100 rs1-wrap-input100 validate-input m-b-20" data-validate="Type user name">
                        <input id="first-name" class="input100" type="text" name="username" placeholder="User name">
                        <span class="focus-input100"></span>
                    </div>
                    <div class="wrap-input100 rs2-wrap-input100 validate-input m-b-20" data-validate="Type password">
                        <input class="input100" type="password" name="pass" placeholder="Password">
                        <span class="focus-input100"></span>
                    </div> --}}
                    <div class="container-login100-form-btn">
                        <button class="login100-form-btn" type="submit">
                            Download
                        </button>
                    </div>
                    <div class="w-full text-center p-t-27 p-b-239">
                        {{-- <span class="txt1">
                        Forgot
                        </span>
                        <a href="/#" class="txt2">
                        User name / password?
                        </a>
                        </div>
                        <div class="w-full text-center">
                        <a href="/#" class="txt3">
                        Sign Up
                        </a> --}}
                    </div>
                </form>
                <div class="login100-more" style="background-image: url('images/bg-01.jpg');"></div>
            </div>
        </div>
    </div>
    <div id="dropDownSelect1"></div>

    <script src="/vendor/jquery/jquery-3.2.1.min.js"></script>

    <script src="/vendor/animsition/js/animsition.min.js"></script>

    <script src="/vendor/bootstrap/js/popper.js"></script>
    <script src="/vendor/bootstrap/js/bootstrap.min.js"></script>

    <script src="/vendor/select2/select2.min.js"></script>
    <script>
        $(".selection-2").select2({
            minimumResultsForSearch: 20,
            dropdownParent: $('#dropDownSelect1')
        });

    </script>

    <script src="/vendor/daterangepicker/moment.min.js"></script>
    <script src="/vendor/daterangepicker/daterangepicker.js"></script>

    <script src="/vendor/countdowntime/countdowntime.js"></script>

    <script src="/js/main.js"></script>

    <script>
        $(function(){
            $("#year_to").change(function(){
                if($("#year_from").val() > $(this).val()){
                    $("#year_from").val($(this).val());
                }
                $("#year_from").prop('max', $(this).val());
            });
        });
    </script>

    <!-- Mirrored from colorlib.com/etc/lf/Login_v17/index.html by HTTrack Website Copier/3.x [XR&CO'2014], Tue, 23 Feb 2021 04:21:39 GMT -->

</html>
