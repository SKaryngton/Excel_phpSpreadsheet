<?php

namespace App\Controller;

use App\service\PhpSpreadsheet;
use Psr\Log\LoggerInterface;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Annotation\Route;

class HomeController extends AbstractController
{

    #[Route('/', name: 'app_home')]
    public function index(): Response
    {
        return $this->render('home/index.html.twig');
    }

    #[Route('/create', name: 'app_create')]
    public function create(PhpSpreadsheet $spreadsheet, LoggerInterface $logger): Response
    {
        //$spreadsheet->generateAndDownloadSimpleExcelSheet();
        //$spreadsheet->generateAndSaveOnServerSimpleExcelSheet();
        //$value=$spreadsheet->readASheetCell();
        //$logger->info($value);
        //$spreadsheet->addWorksheet();
        //$spreadsheet->copyWorksheet();
        //$spreadsheet->deleteWorksheet();
        // $spreadsheet->selectRangeOfCellsWorksheet();
        $spreadsheet->formulasWorksheet();
        return $this->redirectToRoute('app_home');
    }
}
